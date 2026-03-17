"""Aba Classificações — UNSPSC."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill
from logger import log


async def fill_unspsc(page: Page, unspsc_code: str) -> None:
    """Preenche o código UNSPSC na aba Classificações.

    Sequência: Aba Classificações → botão UNSPSC → preencher código →
    Pesquisar → selecionar resultado → Selecionar
    """
    log.info(f"Preenchendo UNSPSC: {unspsc_code}")

    # Navegar para aba Classificações
    await safe_click(page, SELECTORS["tab_classificacoes"])
    await page.wait_for_load_state("networkidle")

    # Clicar no botão UNSPSC
    await safe_click(page, SELECTORS["unspsc_btn"])
    await page.wait_for_load_state("networkidle")

    # Preencher código
    await safe_fill(page, SELECTORS["unspsc_input"], str(unspsc_code))

    # Pesquisar
    await safe_click(page, SELECTORS["unspsc_pesquisar_btn"])
    await page.wait_for_load_state("networkidle")

    # Selecionar o primeiro resultado (ícone de seleção na grid)
    unspsc_select_icon = page.locator("input[id$='kSelUNSPSC']").first
    await unspsc_select_icon.click()
    await page.wait_for_load_state("networkidle")

    # Confirmar seleção
    await safe_click(page, SELECTORS["unspsc_selecionar_btn"])
    await page.wait_for_load_state("networkidle")

    log.info(f"UNSPSC {unspsc_code} selecionado")
