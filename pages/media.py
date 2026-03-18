"""Upload de documentos na aba Mídias."""

from pathlib import Path

from playwright.async_api import Page

from config import SELECTORS, DOCUMENTS_DIR
from browser import safe_click, safe_fill
from logger import log


async def upload_documents(page: Page, doc_files: list[str]) -> None:
    """Faz upload de documentos para a aba Mídias.

    Usa Playwright file chooser nativo — resolve o problema do popup Windows
    que o PAD não tratava bem.
    """
    if not doc_files:
        log.info("Nenhum documento para upload")
        return

    log.info(f"Fazendo upload de {len(doc_files)} documento(s)")

    # Navegar para aba Mídias (abre em nova aba no browser)
    pages_before = len(page.context.pages)
    await safe_click(page, SELECTORS["tab_midias"])

    # Aguardar nova aba abrir
    media_page = page
    for _ in range(10):
        await page.wait_for_timeout(1000)
        if len(page.context.pages) > pages_before:
            break

    # Encontrar a aba de mídias pela URL (Midia.aspx)
    for p in page.context.pages:
        if "Midia.aspx" in p.url:
            media_page = p
            await media_page.wait_for_load_state("networkidle")
            break

    if media_page == page:
        log.warning("Aba de Mídias não abriu separadamente — usando página atual")

    for doc_entry in doc_files:
        doc_path = Path(doc_entry)
        if not doc_path.is_absolute():
            doc_path = DOCUMENTS_DIR / doc_entry
        if not doc_path.exists():
            log.warning(f"Documento não encontrado, pulando: {doc_path}")
            continue
        doc_name = doc_path.stem

        log.debug(f"Uploading: {doc_name}")

        # Clicar em "Adicionar Mídia"
        await safe_click(media_page, SELECTORS["media_add_link"])
        await media_page.wait_for_load_state("networkidle")

        # Usar file chooser do Playwright (evita popup Windows)
        async with media_page.expect_file_chooser() as fc_info:
            await media_page.click(SELECTORS["media_file_input"])
        file_chooser = await fc_info.value
        await file_chooser.set_files(str(doc_path))

        await media_page.wait_for_timeout(2000)

        # Preencher título
        await safe_fill(media_page, SELECTORS["media_titulo_input"], doc_name)

        # Salvar
        salvar_btn = media_page.locator(SELECTORS["media_salvar_btn"]).first
        await salvar_btn.click()
        await media_page.wait_for_load_state("networkidle")
        await media_page.wait_for_timeout(2000)

        log.debug(f"Documento '{doc_name}' uploaded")

    # Fechar aba de mídias e voltar à página principal
    if media_page != page:
        fechar_btn = media_page.locator(SELECTORS["media_fechar_btn"])
        if await fechar_btn.count() > 0:
            try:
                await fechar_btn.click()
            except Exception:
                pass  # cmdFechar tem onclick=window.close() — página fecha imediatamente
        if not media_page.is_closed():
            await media_page.close()
        await page.bring_to_front()
    else:
        fechar_btn = media_page.locator(SELECTORS["media_fechar_btn"])
        if await fechar_btn.count() > 0:
            await fechar_btn.click()
            await page.wait_for_timeout(1000)

    log.info(f"Upload de {len(doc_files)} documento(s) concluído")
