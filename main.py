"""Orquestrador principal do RPA Klassmatt MODEC.

Migração do Power Automate Desktop para Python + Playwright.
Faz cadastro em massa no sistema web Klassmatt.
"""

import asyncio
import sys
import time

from config import KLASSMATT_HOME, MAX_RETRIES, RETRY_DELAY_MS, SELECTORS
from browser import (
    launch_browser, navigate_home, retry_action,
    verificar_sessao, fechar_popups, _handle_dialog, hide_overlays,
)
from excel_handler import load_excel, color_row, save_excel, validate_documents, enrich_missing_data
from state import load_progress, mark_item, is_processed
from logger import log

from pages.worklist import navigate_to_worklist
from pages.item import search_and_select_sin, atuar_no_item, criar_item, finalizar_e_remeter, check_item_already_processed
from pages.classifications import fill_unspsc
from pages.fiscal import fill_ncm
from pages.references import fill_reference
from pages.relationships import fill_relationship
from pages.media import upload_documents
from pages.descriptions import validate_sap_description, change_pdm
from pages.attributes import fill_attributes


# ── Cronômetro de etapas ──

class StepTimer:
    """Mede tempo entre etapas para identificar gargalos da plataforma."""

    def __init__(self, sin: str):
        self.sin = sin
        self.start = time.time()
        self.last = self.start
        self.steps: list[tuple[str, float]] = []

    def mark(self, step_name: str) -> None:
        now = time.time()
        elapsed = now - self.last
        self.steps.append((step_name, elapsed))
        log.info(f"  [timer] {step_name}: {elapsed:.1f}s")
        self.last = now

    def total(self) -> float:
        return time.time() - self.start

    def summary(self) -> str:
        lines = [f"SIN {self.sin} — {self.total():.1f}s total"]
        for name, secs in self.steps:
            bar = "█" * int(secs / 0.5)  # 1 bloco = 0.5s
            lines.append(f"  {secs:5.1f}s {bar} {name}")
        return "\n".join(lines)


class KlassmattSessionError(Exception):
    """Erro de sessão do Klassmatt — requer fechar e reabrir browser."""
    pass


async def _check_page_error(page) -> bool:
    """Verifica se a página atual é a tela de erro do Klassmatt."""
    try:
        text = await page.evaluate("() => document.body.innerText.substring(0, 500)")
        return "exce" in text.lower() or "ACESSO" in text
    except Exception:
        return False


async def _voltar_worklist(page) -> None:
    """Volta para a worklist via JS (padrão fix_items.py).

    Mais resiliente que safe_click — funciona mesmo com overlays.
    Se detectar página de erro/exceção, levanta KlassmattSessionError
    para que o caller feche e reabra o browser.
    """
    # Verificar se já estamos na página de erro
    if await _check_page_error(page):
        raise KlassmattSessionError("Página de erro detectada antes de navegar")

    # Navegar para Principal via JS
    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const p = Array.from(links).find(a => a.innerText.trim() === 'Principal');
        if (p) p.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Verificar se caiu em página de erro
    if await _check_page_error(page):
        raise KlassmattSessionError("Página de erro após navegar para Principal")

    # Navegar para Worklist via JS
    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
        if (wl) wl.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Verificar se pesquisar() está disponível (worklist carregada)
    ready = await page.evaluate("() => typeof pesquisar === 'function'")
    if ready:
        await page.select_option(
            SELECTORS["worklist_filter_dropdown"],
            label="Todas as Solicitações",
        )
        await page.evaluate("() => { pesquisar(0, ''); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
        log.info("Worklist filtrada: Todas as Solicitações")
        return

    raise KlassmattSessionError("Worklist não carregou — pesquisar() indisponível")


async def process_item(page, item: dict, wb) -> tuple[str, list[str]]:
    """Processa um único item do Excel.

    Retorna: (status, warnings) onde status é 'ok', 'needs_review',
    'duplicate', 'error', ou 'skipped'.
    """
    sin = str(item["sin"])
    row = item["_row"]
    t = StepTimer(sin)
    warnings: list[str] = []

    log.info(f"{'='*60}")
    log.info(f"Processando SIN {sin} (linha {row})")
    log.info(f"{'='*60}")

    # Fechar popups e esconder overlays
    await fechar_popups(page)

    # 1. Buscar e selecionar o SIN na worklist
    await search_and_select_sin(page, sin)
    t.mark("Buscar SIN")

    # 2. Atuar no Item (detecta botão disabled/APROVACAO-TECNICA antes de clicar)
    skip_status = await atuar_no_item(page)
    if skip_status:
        # Item não editável (botão disabled, status não-FINALIZACAO, etc.)
        log.info(f"Item em '{skip_status}' — pulando")
        t.mark(f"Status: {skip_status}")
        await _voltar_worklist(page)
        t.mark("Voltar Worklist")
        log.info(f"\n{t.summary()}")
        return "skipped", []
    await hide_overlays(page)
    t.mark("Atuar no Item")

    # 2b. Verificar status do item no workflow (fallback pós-navegação)
    item_status = await check_item_already_processed(page)
    if item_status:
        log.info(f"Item em '{item_status}' — pulando")
        t.mark(f"Status: {item_status}")
        await _voltar_worklist(page)
        t.mark("Voltar Worklist")
        log.info(f"\n{t.summary()}")
        return "skipped", []

    # 3. Criar Item → Finalizar → Salvar → Sim
    await criar_item(page)
    t.mark("Criar Item")

    # 4. UNSPSC
    if item.get("unspsc"):
        unspsc_ok = await fill_unspsc(page, str(item["unspsc"]))
        t.mark("UNSPSC")
        if not unspsc_ok:
            warnings.append("unspsc_not_found")

    # 5. Fiscal — NCM
    if item.get("ncm"):
        ncm_ok = await fill_ncm(page, str(item["ncm"]))
        t.mark("NCM")
        if not ncm_ok:
            warnings.append("ncm_rejected")

    # 6. Referências — Empresa + Part Number
    if item.get("empresa") and item.get("part_number"):
        ref_ok = await fill_reference(page, str(item["empresa"]), str(item["part_number"]))
        t.mark("Referências")
        if not ref_ok:
            color_row(wb, row, "duplicate")
            save_excel(wb)
            await navigate_home(page)
            await navigate_to_worklist(page)
            log.info(f"\n{t.summary()}")
            return "duplicate", []
    else:
        missing = []
        if not item.get("empresa"):
            missing.append("empresa")
        if not item.get("part_number"):
            missing.append("part_number")
        log.warning(f"SIN {sin}: pulando Referências — campos vazios: {', '.join(missing)}")

    # 7. Relacionamento — CÓDIGO ANTIGO
    if item.get("codigo_60"):
        await fill_relationship(page, str(item["codigo_60"]))
        t.mark("Relacionamento")
    else:
        log.warning(f"SIN {sin}: pulando Relacionamento — codigo_60 vazio")

    # 8. Upload de documentos
    doc_files = item.get("_doc_files", [])
    if doc_files:
        await upload_documents(page, doc_files)
        t.mark(f"Upload Mídias ({len(doc_files)} docs)")

    # 9. Validar descrição SAP (Exibe D2 / 40 chars)
    await validate_sap_description(page)
    t.mark("Validação SAP")

    # 10. Alterar PDM
    if item.get("pdm"):
        await change_pdm(page, str(item["pdm"]))
        t.mark("Alterar PDM")

    # 11. Preencher atributos técnicos
    attrs_ok = await fill_attributes(page, item.get("attributes", []))
    t.mark("Atributos")
    if not attrs_ok:
        warnings.append("attributes_incomplete")

    # 12. Finalizar e Remeter para MODEC
    # DESABILITADO até validação completa — manter itens em FINALIZACAO
    # await finalizar_e_remeter(page)
    # t.mark("Remeter MODEC")
    log.info("Remeter MODEC DESABILITADO — item mantido em FINALIZACAO para revisão")
    t.mark("(Remeter desabilitado)")

    # Voltar para worklist (mesmo padrão do fix_items.py)
    # Buffer pós-atributos — rate limiting principal é o asyncio.sleep(5) entre itens
    await page.wait_for_timeout(3_000)
    await _voltar_worklist(page)
    t.mark("Voltar Worklist")

    log.info(f"\n{t.summary()}")

    if warnings:
        log.warning(f"SIN {sin}: processado com problemas: {warnings}")
        return "needs_review", warnings
    return "ok", []


async def _restart_browser(pw, context) -> tuple:
    """Fecha browser e reabre. Retorna (context, page)."""
    log.warning("Fechando browser e reabrindo...")
    try:
        await context.close()
    except Exception:
        pass
    try:
        await pw.stop()
    except Exception:
        pass
    # Aguardar processos morrerem
    await asyncio.sleep(5)
    new_pw, new_context, new_page = await launch_browser()
    await navigate_home(new_page)
    sessao_ok = await verificar_sessao(new_page)
    if not sessao_ok:
        log.warning("Sessão expirada após reabrir — aguardando 60s para re-login manual")
        await asyncio.sleep(60)
        await new_page.reload(wait_until="networkidle")
    await _voltar_worklist(new_page)
    return new_pw, new_context, new_page


async def process_item_with_retry(page, context, pw, item: dict, wb, progress: dict) -> tuple:
    """Executa process_item com retry e recuperação de erro.

    Padrão do bot_pso: retry com backoff crescente.
    Retorna (status, page, context, pw) — browser pode ser recriado.
    """
    sin = str(item.get("sin", ""))
    row = item["_row"]

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            status, item_warnings = await process_item(page, item, wb)
            mark_item(progress, sin, status, warnings=item_warnings if item_warnings else None)
            color_row(wb, row, status)
            save_excel(wb)
            return status, page, context, pw

        except KlassmattSessionError as e:
            log.error(f"Erro de sessão Klassmatt SIN {sin}: {e}")
            pw, context, page = await _restart_browser(pw, context)
            if attempt >= MAX_RETRIES:
                mark_item(progress, sin, "error", str(e))
                color_row(wb, row, "error")
                save_excel(wb)
                return "error", page, context, pw
            continue

        except Exception as e:
            log.error(
                f"ERRO ao processar SIN {sin} (tentativa {attempt}/{MAX_RETRIES}): {e}",
                exc_info=True,
            )

            if attempt < MAX_RETRIES:
                delay = (RETRY_DELAY_MS / 1000) * attempt
                log.info(f"Aguardando {delay:.0f}s antes da próxima tentativa...")
                await asyncio.sleep(delay)

                # Recuperar: voltar para worklist
                try:
                    await _voltar_worklist(page)
                except KlassmattSessionError:
                    log.warning("Sessão perdida na recuperação — reiniciando browser")
                    pw, context, page = await _restart_browser(pw, context)
                except Exception as recovery_error:
                    log.error(f"Falha na recuperacao: {recovery_error}")
                    try:
                        pw, context, page = await _restart_browser(pw, context)
                    except Exception:
                        log.error("Falha ao reiniciar browser")
                        raise
            else:
                mark_item(progress, sin, "error", str(e))
                color_row(wb, row, "error")
                save_excel(wb)
                log.error(
                    f"SIN {sin} falhou apos {MAX_RETRIES} tentativas. "
                    f"Continuando com proximo item."
                )
                # Recuperar para o próximo item
                try:
                    await _voltar_worklist(page)
                except Exception:
                    try:
                        pw, context, page = await _restart_browser(pw, context)
                    except Exception:
                        log.error("Falha ao reiniciar browser para próximo item")
                        raise

                return "error", page, context, pw

    return "error", page, context, pw


async def run() -> None:
    """Fluxo principal do RPA."""
    run_start = time.time()
    log.info("=" * 60)
    log.info("RPA Klassmatt MODEC — Início")
    log.info("=" * 60)

    # Carregar Excel
    wb, items = load_excel()

    # Preencher campos vazios usando dados de itens vizinhos
    items = enrich_missing_data(items)

    # Validar documentos antecipadamente
    items = validate_documents(items)

    # Carregar progresso (para retomada)
    progress = load_progress()

    # Iniciar browser
    pw, context, page = await launch_browser()

    item_times: list[float] = []  # tempo de cada item processado

    try:
        # Navegar para home
        await navigate_home(page)

        # Verificar sessão antes de começar
        sessao_ok = await verificar_sessao(page)
        if not sessao_ok:
            log.warning("Sessao nao detectada -- aguardando 60s para login manual no browser")
            await asyncio.sleep(60)
            await page.reload(wait_until="networkidle")

        await navigate_to_worklist(page)
        setup_elapsed = time.time() - run_start
        log.info(f"[timer] Setup (Excel + browser + worklist): {setup_elapsed:.1f}s")

        total = len(items)
        processed = 0
        errors = 0
        skipped = 0
        needs_review = 0

        for idx, item in enumerate(items):
            sin = str(item.get("sin", ""))
            row = item["_row"]

            # Pular itens já processados com sucesso
            if is_processed(progress, sin):
                log.info(f"SIN {sin} já processado — pulando")
                skipped += 1
                continue

            # Pular itens com documentos faltantes
            if item.get("_missing_docs"):
                log.warning(
                    f"SIN {sin}: pulando por documentos faltantes: {item['_missing_docs']}"
                )
                mark_item(progress, sin, "skipped", f"docs faltantes: {item['_missing_docs']}")
                color_row(wb, row, "skipped")
                save_excel(wb)
                skipped += 1
                continue

            # Delay entre itens para não sobrecarregar o Klassmatt
            if idx > 0:
                await asyncio.sleep(5)

            item_start = time.time()
            status, page, context, pw = await process_item_with_retry(page, context, pw, item, wb, progress)
            item_elapsed = time.time() - item_start
            item_times.append(item_elapsed)

            if status == "error":
                errors += 1
            elif status == "needs_review":
                needs_review += 1
            elif status == "skipped":
                skipped += 1
            else:
                processed += 1

            avg_time = sum(item_times) / len(item_times)
            remaining = total - (idx + 1)
            eta = avg_time * remaining

            log.info(
                f"Progresso: {idx + 1}/{total} | OK: {processed} | Revisão: {needs_review} "
                f"| Erros: {errors} | Pulados: {skipped} "
                f"| Item: {item_elapsed:.1f}s | Média: {avg_time:.1f}s | ETA: {eta / 60:.0f}min"
            )

        run_elapsed = time.time() - run_start
        avg = sum(item_times) / len(item_times) if item_times else 0

        log.info("=" * 60)
        log.info(f"RPA concluído! Total: {total} | OK: {processed} | Revisão: {needs_review} | Erros: {errors} | Pulados: {skipped}")
        log.info(f"[timer] Tempo total: {run_elapsed:.0f}s ({run_elapsed / 60:.1f}min) | Média/item: {avg:.1f}s")
        log.info("=" * 60)

    except Exception as e:
        run_elapsed = time.time() - run_start
        log.error(f"Erro fatal apos {run_elapsed:.0f}s: {e}", exc_info=True)
        log.info("Browser mantido aberto para inspecao.")

    finally:
        try:
            await context.close()
        except Exception:
            log.debug("Browser já estava fechado")
        try:
            await pw.stop()  # type: ignore
        except Exception:
            pass


def main():
    """Entry point."""
    asyncio.run(run())


if __name__ == "__main__":
    main()
