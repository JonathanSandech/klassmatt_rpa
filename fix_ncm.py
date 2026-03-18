"""Script para corrigir NCM em itens já processados.

Fluxo por item:
1. Buscar SIN na worklist
2. Se APROVACAO-TECNICA → Retornar Etapa → Sim
3. Atuar no Item → Fiscal → preencher NCM formatado → Salvar
4. Remeter Modec → Sim
"""

import asyncio
import re
import time

from playwright.async_api import async_playwright, Dialog

from config import EXCEL_PATH, PROFILE_DIR, SLOW_MO, HEADLESS, VIEWPORT_WIDTH, VIEWPORT_HEIGHT
from excel_handler import load_excel
from logger import log


def format_ncm(ncm: str) -> str:
    digits = re.sub(r"\D", "", str(ncm))
    if len(digits) == 8:
        return f"{digits[:4]}.{digits[4:6]}.{digits[6:8]}"
    return str(ncm)


async def handle_dialog(dialog: Dialog):
    log.debug(f"Dialog ({dialog.type}): {dialog.message}")
    await dialog.accept()


async def fix_ncm_for_sin(page, sin: str, ncm_formatted: str) -> str:
    """Retorna 'ok', 'skipped', ou 'error'."""
    context = page.context

    log.info(f"--- SIN {sin} | NCM → {ncm_formatted} ---")

    # Garantir que estamos na worklist com pesquisar disponível
    for attempt in range(3):
        ready = await page.evaluate("() => typeof pesquisar === 'function'")
        if ready:
            break
        # Recarregar worklist
        await page.evaluate("""() => {
            const links = document.querySelectorAll('a');
            const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
            const principal = Array.from(links).find(a => a.innerText === 'Principal');
            if (wl) wl.click();
            else if (principal) principal.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)

    # Buscar SIN
    await page.evaluate(f"""() => {{
        const ta = document.querySelector("textarea[name$='txtValor']");
        if (ta) ta.value = '{sin}';
        pesquisar(0, '');
    }}""")
    await page.wait_for_timeout(5000)

    # Abrir SIN
    await page.evaluate(f"() => {{ abreSIN({sin}); }}")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)

    # Checar status
    status = await page.evaluate(
        "() => { const s = document.querySelector(\"input[id$='txtStatus']\"); return s ? s.value : ''; }"
    )
    log.info(f"  Status: {status}")

    if status == "APROVACAO-TECNICA":
        # Retornar Etapa
        await page.evaluate("""() => {
            const btn = document.querySelector('#lkbutTrazerDeVolta');
            if (btn) btn.click();
        }""")
        await page.wait_for_timeout(3000)

        # Clicar Sim no popup de confirmação
        sim_btn = page.locator("input[value='Sim']")
        try:
            await sim_btn.click(timeout=5000)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(3000)
        except Exception:
            log.warning(f"  Botão Sim não encontrado para Retornar Etapa")
            # Voltar para worklist
            await page.evaluate("""() => {
                const links = document.querySelectorAll('a');
                const v = Array.from(links).find(a => a.innerText === 'Voltar');
                if (v) v.click();
            }""")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2000)
            return "error"

        # Re-checar status
        status = await page.evaluate(
            "() => { const s = document.querySelector(\"input[id$='txtStatus']\"); return s ? s.value : ''; }"
        )
        log.info(f"  Status após retorno: {status}")

    if status != "FINALIZACAO":
        log.warning(f"  Status {status} — não é FINALIZACAO, pulando")
        await page.evaluate("""() => {
            const links = document.querySelectorAll('a');
            const v = Array.from(links).find(a => a.innerText === 'Voltar');
            if (v) v.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2000)
        return "skipped"

    # Checar se NCM já está preenchido na página de resumo
    ncm_atual = await page.evaluate(
        """() => {
            const inputs = document.querySelectorAll('input[type="text"]');
            for (const i of inputs) {
                if (i.value && i.value.match(/^\\d{4}\\.\\d{2}\\.\\d{2}$/)) return i.value;
            }
            return '';
        }"""
    )
    if ncm_atual:
        log.info(f"  NCM já preenchido: {ncm_atual} — apenas remetendo")

    # Atuar no Item
    await page.evaluate("""() => {
        const btn = document.querySelector("input[value='Atuar no Item']");
        if (btn) btn.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)

    if not ncm_atual:
        # Clicar aba Fiscal
        await page.evaluate("""() => {
            const tabs = document.querySelectorAll('a');
            const tab = Array.from(tabs).find(a => a.innerText.includes('Fiscal'));
            if (tab) tab.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2000)

        # Verificar se editável
        editable = await page.evaluate(
            "() => { const n = document.querySelector('#txtNCMTIPI'); return n ? !n.readOnly : false; }"
        )
        if not editable:
            log.warning(f"  NCM readonly — pulando")
        else:
            # Preencher NCM e validar
            await page.evaluate(f"""() => {{
                const ncm = document.querySelector('#txtNCMTIPI');
                ncm.value = '{ncm_formatted}';
                ncm.focus();
                getDescricaoNCM('NCM');
            }}""")
            await page.wait_for_timeout(5000)

            # Verificar se ficou
            val = await page.evaluate(
                "() => { const n = document.querySelector('#txtNCMTIPI'); return n ? n.value : ''; }"
            )
            if val == ncm_formatted:
                log.info(f"  NCM {ncm_formatted} aceito!")
                # Salvar
                await page.evaluate("() => { document.querySelector('#butSalvar').click(); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(3000)
            else:
                log.warning(f"  NCM rejeitado (valor atual: {val})")

    # Remeter Modec
    remeter = page.locator("input[value='Remeter Modec']")
    if await remeter.count() > 0:
        await remeter.click()
        await page.wait_for_timeout(3000)

        # Confirmar
        sim_btn = page.locator("input[value='Sim']")
        try:
            await sim_btn.click(timeout=5000)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(3000)
        except Exception:
            log.warning(f"  Confirmação Remeter não encontrada")
    else:
        log.warning(f"  Botão Remeter Modec não encontrado")
        # Voltar
        await page.evaluate("""() => {
            const links = document.querySelectorAll('a');
            const v = Array.from(links).find(a => a.innerText === 'Voltar');
            if (v) v.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2000)

    log.info(f"  SIN {sin} concluído!")
    return "ok"


async def run():
    log.info("=" * 50)
    log.info("Fix NCM — Início")
    log.info("=" * 50)

    # Ler Excel para pegar SIN→NCM
    wb, items = load_excel()
    sin_ncm = {}
    for item in items:
        sin = str(item.get("sin", ""))
        ncm = str(item.get("ncm", ""))
        if sin and ncm:
            sin_ncm[sin] = format_ncm(ncm)
    log.info(f"Total de itens: {len(sin_ncm)}")

    # Itens já feitos (atualizar conforme progresso)
    done = {"474677", "474728", "474306", "474470", "474308", "474323", "474335",
            "473417", "473413", "473435", "473436", "473995", "473971", "473967"}

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
    await page.goto("https://modec.klassmatt.com.br/MenuPrincipal.aspx", wait_until="networkidle")
    await page.click("text=Acompanhamento das Solicitações (Worklist)")
    await page.wait_for_load_state("networkidle")
    await page.select_option("select:has(option[value='SOMENTE_REC_ACAO'])", label="Todas as Solicitações")
    await page.evaluate("() => { pesquisar(0, ''); }")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)

    results = {"ok": 0, "skipped": 0, "error": 0}

    for sin, ncm in sin_ncm.items():
        if sin in done:
            log.info(f"SIN {sin} já corrigido — pulando")
            results["ok"] += 1
            continue

        try:
            result = await fix_ncm_for_sin(page, sin, ncm)
            results[result] = results.get(result, 0) + 1
            await page.wait_for_timeout(3000)  # delay entre itens
        except Exception as e:
            log.error(f"Erro SIN {sin}: {e}")
            results["error"] += 1
            # Tentar recuperar
            try:
                await page.goto("https://modec.klassmatt.com.br/MenuPrincipal.aspx", wait_until="networkidle")
                await page.click("text=Acompanhamento das Solicitações (Worklist)")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(3000)
                await page.select_option("select:has(option[value='SOMENTE_REC_ACAO'])", label="Todas as Solicitações")
                await page.evaluate("() => { pesquisar(0, ''); }")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(3000)
            except Exception:
                log.error("Falha na recuperação — encerrando")
                break

    log.info("=" * 50)
    log.info(f"Fix NCM concluído! OK: {results['ok']} | Pulados: {results['skipped']} | Erros: {results['error']}")
    log.info("=" * 50)

    await context.close()
    await pw.stop()


if __name__ == "__main__":
    asyncio.run(run())
