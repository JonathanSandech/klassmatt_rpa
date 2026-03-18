"""Aba Relacionamentos — CÓDIGO ANTIGO, ATIVO ERP, ZBRA."""

from playwright.async_api import Page

from config import SELECTORS, RELATIONSHIP_TYPE, RELATIONSHIP_STATUS, RELATIONSHIP_COMMENT
from browser import safe_click, safe_fill
from logger import log


async def fill_relationship(page: Page, codigo_60: str) -> None:
    """Preenche relacionamento com código antigo.

    Tipo: CÓDIGO ANTIGO
    Código: valor de 'Código 60' do Excel
    Status: ATIVO ERP
    Comentário: ZBRA
    """
    log.info(f"Preenchendo relacionamento: {codigo_60}")

    # Navegar para aba Relacionamentos
    await safe_click(page, SELECTORS["tab_relacionamentos"])
    await page.wait_for_load_state("networkidle")

    # Clicar no botão de adicionar
    add_btn = page.locator(SELECTORS["rel_add_btn"])
    await add_btn.click()
    await page.wait_for_load_state("networkidle")

    # Selecionar tipo: CÓDIGO ANTIGO
    await safe_click(page, SELECTORS["rel_tipo_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_TYPE}')")

    # Preencher código
    await safe_fill(page, SELECTORS["rel_codigo_input"], str(codigo_60))

    # Selecionar status: ATIVO ERP
    await safe_click(page, SELECTORS["rel_status_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_STATUS}')")

    # Preencher comentário: ZBRA
    await safe_fill(page, SELECTORS["rel_comentario_input"], RELATIONSHIP_COMMENT)

    # Salvar relacionamento
    save_btn = page.locator(SELECTORS["rel_save_btn"])
    await save_btn.click()
    await page.wait_for_load_state("networkidle")

    log.info(f"Relacionamento salvo: {RELATIONSHIP_TYPE} / {codigo_60}")
