"""Upload de documentos na aba Mídias.

Resilience notes (2026-03):
    - "Adicionar Mídia" é um link __doPostBack dentro de um DataList (dlMidias).
    - Após uploads consecutivos (5+ docs), o ASP.NET re-render pode falhar
      e o link desaparece do DOM. Retry 0% eficaz sem reabrir a aba.
    - Fix: safe_click com timeout curto → fallback __doPostBack direto →
      fallback fechar/reabrir aba Mídias.
"""

import re
from pathlib import Path

from playwright.async_api import Page

from config import SELECTORS, DOCUMENTS_DIR
from browser import safe_click, safe_fill
from logger import log


async def _open_media_tab(page: Page) -> Page:
    """Abre a aba Mídias (Midia.aspx) do ITEM e retorna a page da nova aba.

    Fecha abas de mídias residuais (tipo=SIN) antes de abrir a do item.
    """
    # Fechar abas de mídias residuais (podem ser da SIN, não do item)
    for p in page.context.pages:
        if p != page and "Midia.aspx" in p.url:
            try:
                await p.close()
            except Exception:
                pass

    pages_before = len(page.context.pages)
    await safe_click(page, SELECTORS["tab_midias"])

    # Aguardar nova aba abrir
    for _ in range(10):
        await page.wait_for_timeout(1000)
        if len(page.context.pages) > pages_before:
            break

    # Encontrar a aba de mídias do ITEM pela URL (tipo=Itens)
    for p in page.context.pages:
        if "Midia.aspx" in p.url and "tipo=Itens" in p.url:
            await p.wait_for_load_state("networkidle")
            return p

    # Fallback: qualquer Midia.aspx nova
    for p in page.context.pages:
        if "Midia.aspx" in p.url and p != page:
            await p.wait_for_load_state("networkidle")
            return p

    log.warning("Aba de Mídias não abriu separadamente — usando página atual")
    return page


async def _close_media_tab(media_page: Page, main_page: Page) -> None:
    """Fecha a aba Mídias e volta pra página principal."""
    if media_page == main_page:
        fechar_btn = media_page.locator(SELECTORS["media_fechar_btn"])
        if await fechar_btn.count() > 0:
            await fechar_btn.click()
            await main_page.wait_for_timeout(500)
        return

    fechar_btn = media_page.locator(SELECTORS["media_fechar_btn"])
    if await fechar_btn.count() > 0:
        try:
            await fechar_btn.click()
        except Exception:
            pass  # cmdFechar tem onclick=window.close()
    if not media_page.is_closed():
        await media_page.close()
    await main_page.bring_to_front()


async def _get_existing_docs(media_page: Page) -> dict:
    """Retorna contagem e nomes dos documentos já existentes na aba Mídias."""
    return await media_page.evaluate(
        """() => {
            const names = [];
            const allText = document.body.innerText || '';

            // 1. Contar pelo label "PDF (N)" no texto
            const pdfMatch = allText.match(/PDF\\s*\\((\\d+)\\)/);
            let count = pdfMatch ? parseInt(pdfMatch[1]) : 0;

            // 2. Links de download de mídia (GetMidia = doc real, não UI chrome)
            const mediaLinks = document.querySelectorAll('a[href*="GetMidia"], a[href*="getMidia"]');
            if (mediaLinks.length > count) count = mediaLinks.length;

            // 3. Nomes de documentos (padrão XXXX-XX-XXXX-*)
            const spans = document.querySelectorAll('span, td, div, a');
            for (const s of spans) {
                const text = (s.innerText || s.textContent || '').trim();
                if (text.match(/^\\d{4}-\\d{2}-\\d{4}/) || text.match(/^\\d{4}-[A-Z].*\\.pdf$/i)) {
                    names.push(text.replace(/\\.pdf$/i, ''));
                }
            }
            const uniqueNames = [...new Set(names)];
            if (uniqueNames.length > count) count = uniqueNames.length;

            return { count, names: uniqueNames };
        }"""
    )


async def _click_adicionar_midia(media_page: Page) -> bool:
    """Clica em "Adicionar Mídia" com fallback resiliente.

    1. safe_click com timeout curto (5s)
    2. Fallback: extrair __doPostBack target do link e chamar direto via JS
    3. Se ambos falharem, retorna False (caller deve reabrir a aba)
    """
    # Tentativa 1: safe_click com timeout curto
    try:
        await safe_click(media_page, SELECTORS["media_add_link"], timeout=5000)
        await media_page.wait_for_load_state("networkidle")
        return True
    except Exception:
        log.debug("safe_click 'Adicionar Mídia' falhou — tentando __doPostBack direto")

    # Tentativa 2: extrair target do link e chamar __doPostBack direto
    try:
        clicked = await media_page.evaluate(
            """() => {
                // Buscar link "Adicionar Mídia" e extrair __doPostBack target
                const links = document.querySelectorAll('a');
                const addLink = Array.from(links).find(a => a.innerText.includes('Adicionar'));
                if (addLink) {
                    const href = addLink.getAttribute('href') || '';
                    const match = href.match(/__doPostBack\\('([^']+)'/);
                    if (match) {
                        __doPostBack(match[1], '');
                        return 'postback';
                    }
                    // Fallback: click direto no link
                    addLink.click();
                    return 'click';
                }
                // Último recurso: tentar o target padrão do DataList
                if (typeof __doPostBack === 'function') {
                    // Buscar qualquer Linkbutton1 no dlMidias
                    const btn = document.querySelector("a[id='Linkbutton1']");
                    if (btn) {
                        const href = btn.getAttribute('href') || '';
                        const match = href.match(/__doPostBack\\('([^']+)'/);
                        if (match) {
                            __doPostBack(match[1], '');
                            return 'postback-fallback';
                        }
                    }
                }
                return null;
            }"""
        )
        if clicked:
            log.debug(f"Adicionar Mídia via JS: {clicked}")
            await media_page.wait_for_load_state("networkidle")
            return True
    except Exception as e:
        log.debug(f"__doPostBack fallback falhou: {e}")

    return False


async def upload_documents(page: Page, doc_files: list[str]) -> None:
    """Faz upload de documentos para a aba Mídias.

    Usa Playwright file chooser nativo. Se "Adicionar Mídia" desaparecer
    após uploads consecutivos, fecha e reabre a aba Mídias.
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

    # Abrir aba Mídias
    media_page = await _open_media_tab(page)

    # Verificar documentos já existentes DENTRO da aba de mídias
    existing_docs = await _get_existing_docs(media_page)
    existing_count = existing_docs.get("count", 0)
    existing_names = existing_docs.get("names", [])

    if existing_count >= expected_count:
        log.info(
            f"Mídias já existem na aba ({existing_count} PDFs >= {expected_count} esperados) — pulando. "
            f"Docs: {existing_names}"
        )
        await _close_media_tab(media_page, page)
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

        already_uploaded = any(doc_name in existing for existing in existing_names)
        if already_uploaded:
            log.debug(f"Documento '{doc_name}' já existe na aba de mídias — pulando")
            continue

        docs_to_upload.append((doc_path, doc_name))

    if not docs_to_upload:
        log.info("Todos os documentos já foram uploaded — nada a fazer")
        await _close_media_tab(media_page, page)
        return

    for i, (doc_path, doc_name) in enumerate(docs_to_upload):
        log.debug(f"Uploading ({i+1}/{len(docs_to_upload)}): {doc_name}")

        # Clicar em "Adicionar Mídia" (com fallback resiliente)
        add_ok = await _click_adicionar_midia(media_page)

        if not add_ok:
            # Último recurso: fechar e reabrir aba Mídias pra forçar re-render
            log.warning("Adicionar Mídia não disponível — reabrindo aba Mídias")
            await _close_media_tab(media_page, page)
            await page.wait_for_timeout(2000)
            media_page = await _open_media_tab(page)

            # Tentar de novo após reabrir
            add_ok = await _click_adicionar_midia(media_page)
            if not add_ok:
                log.error(f"Adicionar Mídia falhou mesmo após reabrir aba — abortando upload de '{doc_name}'")
                break

        # Verificar se o form de upload apareceu
        file_input = media_page.locator(SELECTORS["media_file_input"])
        try:
            await file_input.wait_for(state="visible", timeout=5000)
        except Exception:
            log.warning(f"Form de upload não apareceu para '{doc_name}' — pulando")
            continue

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
        await media_page.wait_for_timeout(1000)

        log.debug(f"Documento '{doc_name}' uploaded")

    # Fechar aba de mídias e voltar à página principal
    await _close_media_tab(media_page, page)

    log.info(f"Upload de {len(docs_to_upload)} documento(s) concluído")
