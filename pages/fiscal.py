"""Aba Fiscal — preenchimento do NCM."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill
from logger import log


async def fill_ncm(page: Page, ncm: str) -> None:
    """Navega para aba Fiscal e preenche o campo NCM/TIPI."""
    log.info(f"Preenchendo NCM: {ncm}")

    await safe_click(page, SELECTORS["tab_fiscal"])
    await page.wait_for_load_state("networkidle")

    await safe_fill(page, SELECTORS["ncm_input"], str(ncm))

    log.info(f"NCM {ncm} preenchido")
