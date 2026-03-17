"""Orquestrador principal do RPA Klassmatt MODEC.

Migração do Power Automate Desktop para Python + Playwright.
Faz cadastro em massa no sistema web Klassmatt.
"""

import asyncio
import sys
import time

from config import KLASSMATT_HOME, MAX_RETRIES, RETRY_DELAY_MS
from browser import (
    launch_browser, navigate_home, retry_action,
    verificar_sessao, fechar_popups, _handle_dialog,
)
from excel_handler import load_excel, color_row, save_excel, validate_documents
from state import load_progress, mark_item, is_processed
from logger import log

from pages.worklist import navigate_to_worklist
from pages.item import search_and_select_sin, atuar_no_item, criar_item, finalizar_e_remeter
from pages.classifications import fill_unspsc
from pages.fiscal import fill_ncm
from pages.references import fill_reference
from pages.relationships import fill_relationship
from pages.media import upload_documents
from pages.descriptions import validate_sap_description, change_pdm
from pages.attributes import fill_attributes


async def process_item(page, item: dict, wb) -> str:
    """Processa um único item do Excel.

    Retorna: 'ok', 'duplicate', 'error', ou 'skipped'.
    """
    sin = str(item["sin"])
    row = item["_row"]

    log.info(f"{'='*60}")
    log.info(f"Processando SIN {sin} (linha {row})")
    log.info(f"{'='*60}")

    start = time.time()

    # Fechar popups que possam estar bloqueando
    await fechar_popups(page)

    # 1. Buscar e selecionar o SIN na worklist
    await search_and_select_sin(page, sin)

    # 2. Atuar no Item
    await atuar_no_item(page)

    # 3. Criar Item → Finalizar → Salvar → Sim
    await criar_item(page)

    # 4. UNSPSC
    if item.get("unspsc"):
        await fill_unspsc(page, str(item["unspsc"]))

    # 5. Fiscal — NCM
    if item.get("ncm"):
        await fill_ncm(page, str(item["ncm"]))

    # 6. Referências — Empresa + Part Number
    if item.get("empresa") and item.get("part_number"):
        ref_ok = await fill_reference(page, str(item["empresa"]), str(item["part_number"]))
        if not ref_ok:
            # Referência duplicada — pintar laranja e pular
            color_row(wb, row, "duplicate")
            save_excel(wb)
            await navigate_home(page)
            await navigate_to_worklist(page)
            return "duplicate"

    # 7. Relacionamento — CÓDIGO ANTIGO
    if item.get("codigo_60"):
        await fill_relationship(page, str(item["codigo_60"]))

    # 8. Upload de documentos
    doc_files = item.get("_doc_files", [])
    missing_docs = item.get("_missing_docs", [])
    valid_docs = [d for d in doc_files if d not in missing_docs]
    if valid_docs:
        await upload_documents(page, valid_docs)

    # 9. Validar descrição SAP (Exibe D2 / 40 chars)
    await validate_sap_description(page)

    # 10. Alterar PDM
    if item.get("pdm"):
        await change_pdm(page, str(item["pdm"]))

    # 11. Preencher atributos técnicos
    await fill_attributes(page, item.get("attributes", []))

    # 12. Finalizar e Remeter para MODEC
    await finalizar_e_remeter(page)

    elapsed = time.time() - start
    log.info(f"SIN {sin} concluído em {elapsed:.1f}s")

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
                        log.warning("Sessão expirada — aguardando login manual...")
                        await page.wait_for_timeout(30_000)
                    await navigate_to_worklist(page)
                except Exception as recovery_error:
                    log.error(f"Falha na recuperação: {recovery_error}")
                    # Recriar página como fallback
                    page = await context.new_page()
                    page.on("dialog", _handle_dialog)
                    await navigate_home(page)
                    await navigate_to_worklist(page)
            else:
                # Esgotou tentativas
                mark_item(progress, sin, "error", str(e))
                color_row(wb, row, "error")
                save_excel(wb)

                # Recuperar para o próximo item
                try:
                    await navigate_home(page)
                    await navigate_to_worklist(page)
                except Exception as recovery_error:
                    log.error(f"Falha na recuperação final: {recovery_error}")
                    page = await context.new_page()
                    page.on("dialog", _handle_dialog)
                    await navigate_home(page)
                    await navigate_to_worklist(page)

                return "error", page

    return "error", page


async def run() -> None:
    """Fluxo principal do RPA."""
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

    try:
        # Navegar para home
        await navigate_home(page)

        # Verificar sessão antes de começar
        sessao_ok = await verificar_sessao(page)
        if not sessao_ok:
            log.warning(
                "Sessão não detectada — faça login manual no browser e pressione Enter no terminal"
            )
            input("Pressione ENTER após fazer login...")
            await page.reload(wait_until="networkidle")

        await navigate_to_worklist(page)

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

            status, page = await process_item_with_retry(page, context, item, wb, progress)

            if status == "error":
                errors += 1
            else:
                processed += 1

            log.info(
                f"Progresso: {idx + 1}/{total} | OK: {processed} | Erros: {errors} | Pulados: {skipped}"
            )

        log.info("=" * 60)
        log.info(f"RPA concluído! Total: {total} | OK: {processed} | Erros: {errors} | Pulados: {skipped}")
        log.info("=" * 60)

    finally:
        await context.close()
        await pw.stop()  # type: ignore


def main():
    """Entry point."""
    asyncio.run(run())


if __name__ == "__main__":
    main()
