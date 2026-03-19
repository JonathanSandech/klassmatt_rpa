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
from excel_handler import load_excel, color_row, save_excel, validate_documents
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


async def process_item(page, item: dict, wb) -> str:
    """Processa um único item do Excel.

    Retorna: 'ok', 'duplicate', 'error', ou 'skipped'.
    """
    sin = str(item["sin"])
    row = item["_row"]
    t = StepTimer(sin)

    log.info(f"{'='*60}")
    log.info(f"Processando SIN {sin} (linha {row})")
    log.info(f"{'='*60}")

    # Fechar popups e esconder overlays
    await fechar_popups(page)

    # 1. Buscar e selecionar o SIN na worklist
    await search_and_select_sin(page, sin)
    t.mark("Buscar SIN")

    # 2. Atuar no Item
    await atuar_no_item(page)
    await hide_overlays(page)
    t.mark("Atuar no Item")

    # 2b. Verificar se item já foi processado (status != FINALIZACAO)
    if await check_item_already_processed(page):
        log.info(f"\n{t.summary()}")
        return "ok"

    # 3. Criar Item → Finalizar → Salvar → Sim
    await criar_item(page)
    t.mark("Criar Item")

    # 4. UNSPSC
    if item.get("unspsc"):
        await fill_unspsc(page, str(item["unspsc"]))
        t.mark("UNSPSC")

    # 5. Fiscal — NCM
    if item.get("ncm"):
        await fill_ncm(page, str(item["ncm"]))
        t.mark("NCM")

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
            return "duplicate"

    # 7. Relacionamento — CÓDIGO ANTIGO
    if item.get("codigo_60"):
        await fill_relationship(page, str(item["codigo_60"]))
        t.mark("Relacionamento")

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
    await fill_attributes(page, item.get("attributes", []))
    t.mark("Atributos")

    # 12. Finalizar e Remeter para MODEC
    # DESABILITADO até validação completa — manter itens em FINALIZACAO
    # await finalizar_e_remeter(page)
    # t.mark("Remeter MODEC")
    log.info("Remeter MODEC DESABILITADO — item mantido em FINALIZACAO para revisão")
    t.mark("(Remeter desabilitado)")

    # Voltar para worklist
    await navigate_home(page)
    await navigate_to_worklist(page)
    t.mark("Voltar Worklist")

    log.info(f"\n{t.summary()}")

    return "ok"


async def process_item_with_retry(page, context, item: dict, wb, progress: dict) -> tuple:
    """Executa process_item com retry e recuperação de erro.

    Padrão do bot_pso: retry com backoff crescente.
    Retorna (status, page) — page pode ser recriada se houver erro grave.
    """
    sin = str(item.get("sin", ""))
    row = item["_row"]

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            status = await process_item(page, item, wb)
            mark_item(progress, sin, status)
            color_row(wb, row, status)
            save_excel(wb)
            return status, page

        except Exception as e:
            log.error(
                f"ERRO ao processar SIN {sin} (tentativa {attempt}/{MAX_RETRIES}): {e}",
                exc_info=True,
            )

            if attempt < MAX_RETRIES:
                delay = (RETRY_DELAY_MS / 1000) * attempt
                log.info(f"Aguardando {delay:.0f}s antes da próxima tentativa...")
                await asyncio.sleep(delay)

                # Recuperar: voltar para home e worklist
                try:
                    await navigate_home(page)
                    sessao_ok = await verificar_sessao(page)
                    if not sessao_ok:
                        log.warning("Sessao expirada -- aguardando 60s para re-login manual")
                        await asyncio.sleep(60)
                        await page.reload(wait_until="networkidle")
                    await navigate_to_worklist(page)
                except Exception as recovery_error:
                    log.error(f"Falha na recuperacao: {recovery_error}")
                    # Fechar abas extras antes de recriar
                    for p in context.pages[1:]:
                        try:
                            await p.close()
                        except Exception:
                            pass
                    page = await context.new_page()
                    page.on("dialog", _handle_dialog)
                    await navigate_home(page)
                    await navigate_to_worklist(page)
            else:
                # Esgotou tentativas — pausar para inspeção
                mark_item(progress, sin, "error", str(e))
                color_row(wb, row, "error")
                save_excel(wb)

                log.error(
                    f"SIN {sin} falhou apos {MAX_RETRIES} tentativas. "
                    f"Continuando com proximo item."
                )

                # Recuperar para o próximo item
                try:
                    await navigate_home(page)
                    await navigate_to_worklist(page)
                except Exception as recovery_error:
                    log.error(f"Falha na recuperacao final: {recovery_error}")
                    for p in context.pages[1:]:
                        try:
                            await p.close()
                        except Exception:
                            pass
                    page = await context.new_page()
                    page.on("dialog", _handle_dialog)
                    await navigate_home(page)
                    await navigate_to_worklist(page)

                return "error", page

    return "error", page


async def run() -> None:
    """Fluxo principal do RPA."""
    run_start = time.time()
    log.info("=" * 60)
    log.info("RPA Klassmatt MODEC — Início")
    log.info("=" * 60)

    # Carregar Excel
    wb, items = load_excel()

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
            status, page = await process_item_with_retry(page, context, item, wb, progress)
            item_elapsed = time.time() - item_start
            item_times.append(item_elapsed)

            if status == "error":
                errors += 1
            else:
                processed += 1

            avg_time = sum(item_times) / len(item_times)
            remaining = total - (idx + 1)
            eta = avg_time * remaining

            log.info(
                f"Progresso: {idx + 1}/{total} | OK: {processed} | Erros: {errors} | Pulados: {skipped} "
                f"| Item: {item_elapsed:.1f}s | Média: {avg_time:.1f}s | ETA: {eta / 60:.0f}min"
            )

        run_elapsed = time.time() - run_start
        avg = sum(item_times) / len(item_times) if item_times else 0

        log.info("=" * 60)
        log.info(f"RPA concluído! Total: {total} | OK: {processed} | Erros: {errors} | Pulados: {skipped}")
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
