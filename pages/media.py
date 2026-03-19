"""Upload de documentos na aba Mídias."""

import re
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

    expected_count = len(doc_files)

    # Verificar se já existem mídias suficientes pelo label da aba
    tab_el = page.locator("a:has-text('Mídias')").first
    try:
        tab_text = await tab_el.inner_text(timeout=5_000)
        match = re.search(r"\((\d+)\)", tab_text)
        if match:
            existing_count = int(match.group(1))
            if existing_count >= expected_count:
                log.info(f"Mídias já existem ({existing_count} uploads >= {expected_count} esperados) — pulando")
                return
    except Exception:
        pass

    log.info(f"Fazendo upload de {expected_count} documento(s)")

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

    # Verificar documentos já existentes DENTRO da aba de mídias (mais confiável que o label)
    existing_docs = await media_page.evaluate(
        """() => {
            // Procurar thumbnails de PDFs ou nomes de documentos existentes
            const imgs = document.querySelectorAll('img[alt], a[href*="Midia"]');
            const names = [];
            // Procurar textos com nomes de documento abaixo dos thumbnails
            const allText = document.body.innerText;
            const pdfMatch = allText.match(/PDF\\s*\\((\\d+)\\)/);
            const count = pdfMatch ? parseInt(pdfMatch[1]) : 0;
            // Pegar nomes dos documentos existentes
            const spans = document.querySelectorAll('span, td, div');
            for (const s of spans) {
                const text = s.innerText.trim();
                // Padrão de nome de documento: XXXX-XX-XXXX-...
                if (text.match(/^\\d{4}-\\d{2}-\\d{4}/)) {
                    names.push(text);
                }
            }
            return { count, names };
        }"""
    )

    existing_count = existing_docs.get("count", 0)
    existing_names = existing_docs.get("names", [])

    if existing_count >= expected_count:
        log.info(
            f"Mídias já existem na aba ({existing_count} PDFs >= {expected_count} esperados) — pulando. "
            f"Docs: {existing_names}"
        )
        # Fechar aba de mídias
        if media_page != page:
            try:
                await media_page.close()
            except Exception:
                pass
            await page.bring_to_front()
        return

    # Calcular quantos docs faltam
    docs_to_upload = []
    for doc_entry in doc_files:
        doc_path = Path(doc_entry)
        if not doc_path.is_absolute():
            doc_path = DOCUMENTS_DIR / doc_entry
        if not doc_path.exists():
            log.warning(f"Documento não encontrado, pulando: {doc_path}")
            continue
        doc_name = doc_path.stem

        # Verificar se este documento específico já foi uploaded
        already_uploaded = any(doc_name in existing for existing in existing_names)
        if already_uploaded:
            log.debug(f"Documento '{doc_name}' já existe na aba de mídias — pulando")
            continue

        docs_to_upload.append((doc_path, doc_name))

    if not docs_to_upload:
        log.info("Todos os documentos já foram uploaded — nada a fazer")
        if media_page != page:
            try:
                await media_page.close()
            except Exception:
                pass
            await page.bring_to_front()
        return

    for doc_path, doc_name in docs_to_upload:
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

    log.info(f"Upload de {len(docs_to_upload)} documento(s) concluído")
