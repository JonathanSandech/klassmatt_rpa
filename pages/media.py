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

    # Navegar para aba Mídias (abre em nova aba/janela no Klassmatt)
    await safe_click(page, SELECTORS["tab_midias"])
    await page.wait_for_load_state("networkidle")

    # Pegar a página de mídias (pode abrir em popup)
    # O PAD fazia AttachToChromeByTitle com "Klassmatt - Descriptive Standard System"
    media_page = page
    all_pages = page.context.pages
    for p in all_pages:
        title = await p.title()
        if "Descriptive Standard System" in title:
            media_page = p
            break

    for doc_name in doc_files:
        doc_path = DOCUMENTS_DIR / doc_name
        if not doc_path.exists():
            log.warning(f"Documento não encontrado, pulando: {doc_path}")
            continue

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

    # Fechar janela de mídias
    fechar_btn = media_page.locator(SELECTORS["media_fechar_btn"])
    if await fechar_btn.count() > 0:
        await fechar_btn.click()
        await page.wait_for_timeout(1000)

    log.info(f"Upload de {len(doc_files)} documento(s) concluído")
