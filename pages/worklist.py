"""Navegação na Worklist do Klassmatt."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, wait_for_text
from logger import log


async def navigate_to_worklist(page: Page) -> None:
    """Navega para a Worklist e seleciona 'Todas as Solicitações'."""
    log.info("Navegando para Worklist...")

    # Clicar no link da Worklist
    await safe_click(page, SELECTORS["worklist_link"])
    await page.wait_for_load_state("networkidle")

    # Abrir dropdown de filtro e selecionar "Todas as Solicitações"
    await safe_click(page, SELECTORS["worklist_filter_dropdown"])
    await page.wait_for_timeout(500)

    # Preencher busca no dropdown
    search_input = page.locator(SELECTORS["worklist_filter_search"])
    await search_input.fill("Todas as Solicitações")
    await page.wait_for_timeout(1000)

    # Selecionar opção (3x Down + Enter, como no PAD)
    await page.keyboard.press("ArrowDown")
    await page.keyboard.press("ArrowDown")
    await page.keyboard.press("ArrowDown")
    await page.keyboard.press("Enter")

    await page.wait_for_load_state("networkidle")
    log.info("Worklist filtrada: Todas as Solicitações")
