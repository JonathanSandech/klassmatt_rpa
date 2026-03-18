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

    # Selecionar "Todas as Solicitações" via select nativo + disparar onchange
    # O dropdown é um select2 que wrappa o <select> nativo.
    # select_option muda o valor mas não dispara o onchange do select2,
    # então forçamos via JS.
    await page.select_option(
        SELECTORS["worklist_filter_dropdown"],
        label="Todas as Solicitações",
    )
    await page.evaluate("() => { pesquisar(0, ''); }")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)  # Aguardar resultados carregarem
    log.info("Worklist filtrada: Todas as Solicitações")
