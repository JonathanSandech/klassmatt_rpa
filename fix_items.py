"""Script para corrigir itens com dados incorretos ou faltantes.

Reprocessa SINs específicos usando a mesma lógica do main.py,
com suporte a itens em APROVACAO-TECNICA (Retornar Etapa → editar → Remeter).

Uso:
    python fix_items.py                    # corrige todos os SINs listados em SINS_TO_FIX
    python fix_items.py 474284 474291      # corrige SINs específicos via CLI
"""

import asyncio
import sys
import time

from playwright.async_api import async_playwright, Dialog

from config import (
    EXCEL_PATH, PROFILE_DIR, SLOW_MO, HEADLESS,
    VIEWPORT_WIDTH, VIEWPORT_HEIGHT, SELECTORS,
)
from excel_handler import load_excel
from browser import hide_overlays
from logger import log

# Page objects
from pages.classifications import fill_unspsc
from pages.fiscal import fill_ncm
from pages.references import fill_reference
from pages.relationships import fill_relationship
from pages.media import upload_documents
from pages.descriptions import validate_sap_description, change_pdm
from pages.attributes import fill_attributes

# ── SINs a corrigir (atualizar conforme necessário) ──
SINS_TO_FIX = [
    # APROVACAO-TECNICA — precisam Retornar Etapa
    "474284", "474291", "474329", "474001",
    # Dirty form bloqueou PDM/atributos
    "474018", "474010",
    # Atributo 3 existente=N/A mas planilha tem valor
    "473462",
    # Popup árvore falhou (atributo 1)
    "473919",
]

# Se True, remete o item de volta para MODEC após corrigir
REMETER_APOS_FIX = False


async def handle_dialog(dialog: Dialog):
    log.debug(f"Dialog ({dialog.type}): {dialog.message}")
    await dialog.accept()


async def _navigate_to_worklist(page):
    """Garante que estamos na worklist com pesquisar() disponível."""
    for attempt in range(3):
        ready = await page.evaluate("() => typeof pesquisar === 'function'")
        if ready:
            return
        await page.evaluate("""() => {
            const links = document.querySelectorAll('a');
            const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
            const principal = Array.from(links).find(a => a.innerText === 'Principal');
            if (wl) wl.click();
            else if (principal) principal.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)


async def _search_and_open_sin(page, sin: str):
    """Busca SIN na worklist e abre a página do item.

    O link da worklist usa OpenNewTab() que abre SIN_Item_Resultante em nova aba.
    Retorna a page da nova aba (ou a mesma page se navegou inline).
    """
    await page.evaluate(f"""() => {{
        const ta = document.querySelector("textarea[name$='txtValor']");
        if (ta) ta.value = '{sin}';
        pesquisar(0, '');
    }}""")
    await page.wait_for_timeout(5000)

    # Clicar no link do resultado — usa OpenNewTab, abre em nova aba
    pages_before = len(page.context.pages)
    found = await page.evaluate(f"""() => {{
        const results = document.querySelectorAll('#DIVResultado .result a');
        if (results.length > 0) {{ results[0].click(); return true; }}
        return false;
    }}""")
    if not found:
        raise RuntimeError(f"SIN {sin} não encontrado na worklist")

    # Esperar a nova aba abrir
    await page.wait_for_timeout(3000)

    # Retornar a nova aba se abriu
    if len(page.context.pages) > pages_before:
        new_page = page.context.pages[-1]
        new_page.on("dialog", handle_dialog)  # Registrar dialog handler na nova aba
        await new_page.wait_for_load_state("networkidle")
        await new_page.wait_for_timeout(1000)
        return new_page

    # Fallback: navegou na mesma aba
    await page.wait_for_load_state("networkidle")
    return page


async def _get_status(page) -> str:
    """Lê o status atual do item."""
    return await page.evaluate(
        "() => { const s = document.querySelector(\"input[id$='txtStatus']\"); return s ? s.value : ''; }"
    )


async def _retornar_etapa(page) -> bool:
    """Retorna o item de APROVACAO-TECNICA para FINALIZACAO.

    Retorna True se retornou com sucesso, False se falhou.
    """
    log.info("  Retornando etapa (APROVACAO-TECNICA → FINALIZACAO)...")

    await page.evaluate("""() => {
        const btn = document.querySelector('#lkbutTrazerDeVolta');
        if (btn) btn.click();
    }""")
    await page.wait_for_timeout(3000)

    # Clicar Sim no painel de confirmação (inline, não JS dialog)
    sim_btn = page.locator("input[value='Sim']")
    try:
        await sim_btn.click(timeout=5000)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception:
        log.warning("  Botão 'Sim' não encontrado para Retornar Etapa")
        return False

    status = await _get_status(page)
    log.info(f"  Status após retorno: {status}")
    return status == "FINALIZACAO"


async def _atuar_no_item(page):
    """Clica em 'Atuar no Item'."""
    await page.evaluate("""() => {
        const btn = document.querySelector("input[value='Atuar no Item']");
        if (btn) btn.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)
    await hide_overlays(page)


async def _remeter_modec(page):
    """Remete o item para MODEC (Remeter → Sim)."""
    remeter = page.locator("input[value='Remeter Modec']")
    if await remeter.count() > 0:
        await remeter.click()
        await page.wait_for_timeout(3000)

        sim_btn = page.locator("input[value='Sim']")
        try:
            await sim_btn.click(timeout=5000)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(3000)
            log.info("  Item remetido para MODEC")
        except Exception:
            log.warning("  Confirmação Remeter não encontrada")
    else:
        log.warning("  Botão 'Remeter Modec' não encontrado")


async def _voltar_worklist(page):
    """Volta para a worklist navegando via Principal."""
    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const p = Array.from(links).find(a => a.innerText.trim() === 'Principal');
        if (p) p.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
        if (wl) wl.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Filtrar "Todas as Solicitações"
    try:
        await page.select_option(
            "select:has(option[value='SOMENTE_REC_ACAO'])",
            label="Todas as Solicitações",
        )
        await page.evaluate("() => { pesquisar(0, ''); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception:
        pass


async def fix_sin(page, sin: str, item: dict) -> str:
    """Corrige um SIN específico. Retorna 'ok', 'skipped', ou 'error'."""
    log.info(f"{'='*50}")
    log.info(f"Corrigindo SIN {sin}")
    log.info(f"{'='*50}")

    start = time.time()
    worklist_page = page
    item_page = page  # será atualizado se abrir nova aba

    try:
        await _navigate_to_worklist(page)

        # Abrir SIN — retorna a aba onde o item está
        item_page = await _search_and_open_sin(page, sin)
        log.debug(f"  Aba do item: {item_page.url}")

        # Checar status
        status = await _get_status(item_page)
        log.info(f"  Status atual: {status}")

        if status == "APROVACAO-TECNICA":
            ok = await _retornar_etapa(item_page)
            if not ok:
                log.error(f"  Falha ao retornar etapa — pulando SIN {sin}")
                return "error"

        # Atuar no Item — navega na mesma aba para ITEM_Edita.aspx
        # Pode gerar confirm dialog ("outro usuário atuando") que o handler aceita
        atuar_btn = item_page.locator("input[value='Atuar no Item']")
        if await atuar_btn.count() > 0:
            try:
                await atuar_btn.click(timeout=15_000)
            except Exception:
                # Dialog pode ter bloqueado — tentar JS click
                await item_page.evaluate("""() => {
                    const btn = document.querySelector("input[value='Atuar no Item']");
                    if (btn) btn.click();
                }""")
            await item_page.wait_for_load_state("networkidle")
            await item_page.wait_for_timeout(2000)
        else:
            log.warning("  Botão 'Atuar no Item' não encontrado")

        await hide_overlays(item_page)
        log.debug(f"  Página de edição: {item_page.url}")

        # Re-checar status após Atuar
        status = await _get_status(item_page)
        if status and status != "FINALIZACAO":
            log.warning(f"  Status '{status}' após Atuar — não editável, pulando")
            return "skipped"

        # Reprocessar todas as etapas (cada uma é idempotente)
        # UNSPSC
        if item.get("unspsc"):
            await fill_unspsc(item_page, str(item["unspsc"]))
            log.info(f"  UNSPSC: {item['unspsc']}")

        # NCM
        if item.get("ncm"):
            await fill_ncm(item_page, str(item["ncm"]))
            log.info(f"  NCM: {item['ncm']}")

        # Referências
        if item.get("empresa") and item.get("part_number"):
            await fill_reference(item_page, str(item["empresa"]), str(item["part_number"]))
            log.info(f"  Referência: {item['empresa']} / {item['part_number']}")

        # Relacionamento
        if item.get("codigo_60"):
            await fill_relationship(item_page, str(item["codigo_60"]))
            log.info(f"  Relacionamento: {item['codigo_60']}")

        # Upload de documentos
        doc_files = item.get("_doc_files", [])
        if doc_files:
            await upload_documents(item_page, doc_files)
            log.info(f"  Mídias: {len(doc_files)} doc(s)")

        # Validação SAP
        await validate_sap_description(item_page)

        # PDM
        if item.get("pdm"):
            await change_pdm(item_page, str(item["pdm"]))
            log.info(f"  PDM: {item['pdm']}")

        # Atributos
        await fill_attributes(item_page, item.get("attributes", []))

        # Remeter
        if REMETER_APOS_FIX:
            await _remeter_modec(item_page)
        else:
            log.info("  Remeter DESABILITADO — item mantido em FINALIZACAO")

    except Exception as e:
        log.error(f"  Erro ao reprocessar: {e}", exc_info=True)
        elapsed = time.time() - start
        log.info(f"  Tempo: {elapsed:.1f}s")
        return "error"

    finally:
        # Fechar aba extra aberta pelo Klassmatt
        if item_page != worklist_page:
            try:
                if not item_page.is_closed():
                    await item_page.close()
            except Exception:
                pass

    elapsed = time.time() - start
    log.info(f"  SIN {sin} corrigido em {elapsed:.1f}s")
    return "ok"


async def run():
    log.info("=" * 50)
    log.info("Fix Items — Início")
    log.info("=" * 50)

    # Determinar quais SINs corrigir
    if len(sys.argv) > 1:
        sins_to_fix = sys.argv[1:]
    else:
        sins_to_fix = SINS_TO_FIX

    log.info(f"SINs a corrigir: {sins_to_fix}")

    # Ler Excel para pegar dados completos
    wb, items = load_excel()
    sin_data = {}
    for item in items:
        sin = str(item.get("sin", ""))
        if sin in sins_to_fix:
            sin_data[sin] = item

    missing = set(sins_to_fix) - set(sin_data.keys())
    if missing:
        log.warning(f"SINs não encontrados na planilha: {missing}")

    log.info(f"SINs encontrados na planilha: {len(sin_data)}/{len(sins_to_fix)}")

    # Browser
    pw = await async_playwright().start()
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=HEADLESS,
        slow_mo=SLOW_MO,
        viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
        args=[f"--window-size={VIEWPORT_WIDTH},{VIEWPORT_HEIGHT}"],
    )
    context.set_default_timeout(30_000)
    context.set_default_navigation_timeout(60_000)
    page = context.pages[0] if context.pages else await context.new_page()
    page.on("dialog", handle_dialog)

    # Navegar para worklist
    try:
        await page.goto(
            "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
            wait_until="networkidle",
        )
        await page.click("text=Acompanhamento das Solicitações (Worklist)")
        await page.wait_for_load_state("networkidle")
        await page.select_option(
            "select:has(option[value='SOMENTE_REC_ACAO'])",
            label="Todas as Solicitações",
        )
        await page.evaluate("() => { pesquisar(0, ''); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception as e:
        log.error(f"Falha ao inicializar worklist: {e}")
        log.info("Aguardando 60s para login manual...")
        await asyncio.sleep(60)

    results = {"ok": 0, "skipped": 0, "error": 0}

    for sin in sins_to_fix:
        if sin not in sin_data:
            log.warning(f"SIN {sin} não encontrado na planilha — pulando")
            results["skipped"] += 1
            continue

        try:
            result = await fix_sin(page, sin, sin_data[sin])
            results[result] = results.get(result, 0) + 1
            # Voltar para worklist para o próximo
            await _voltar_worklist(page)
            await asyncio.sleep(5)  # delay entre itens
        except Exception as e:
            log.error(f"Erro fatal SIN {sin}: {e}", exc_info=True)
            results["error"] += 1
            # Tentar recuperar
            try:
                await _voltar_worklist(page)
            except Exception:
                log.error("Falha na recuperação — tentando recriar página")
                try:
                    for p in context.pages[1:]:
                        await p.close()
                    page = await context.new_page()
                    page.on("dialog", handle_dialog)
                    await page.goto(
                        "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
                        wait_until="networkidle",
                    )
                    await _voltar_worklist(page)
                except Exception:
                    log.error("Recuperação falhou — encerrando")
                    break

    log.info("=" * 50)
    log.info(
        f"Fix Items concluído! "
        f"OK: {results['ok']} | Pulados: {results['skipped']} | Erros: {results['error']}"
    )
    log.info("=" * 50)

    await context.close()
    await pw.stop()


if __name__ == "__main__":
    asyncio.run(run())
