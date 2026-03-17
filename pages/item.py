"""Busca, criação e finalização de item no Klassmatt."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, wait_for_text
from logger import log


async def search_and_select_sin(page: Page, sin: str) -> None:
    """Busca um SIN na worklist e seleciona o item encontrado."""
    log.info(f"Buscando SIN: {sin}")

    # Preencher campo de busca
    await safe_fill(page, SELECTORS["sin_search"], str(sin))

    # Clicar em Filtrar
    await safe_click(page, SELECTORS["sin_filter_btn"])
    await page.wait_for_load_state("networkidle")

    # Selecionar o primeiro resultado na tabela
    # O PAD clica no span com texto do item — aqui selecionamos a primeira linha de resultado
    result_row = page.locator("table.GridClass tr.GridItemClass, table.GridClass tr.GridAlternateItemClass").first
    await result_row.click()
    await page.wait_for_load_state("networkidle")

    log.info(f"SIN {sin} selecionado")


async def atuar_no_item(page: Page) -> None:
    """Clica em 'Atuar no Item'."""
    await safe_click(page, SELECTORS["atuar_no_item_btn"])
    await page.wait_for_load_state("networkidle")
    log.debug("Clicou em 'Atuar no Item'")


async def criar_item(page: Page) -> None:
    """Cria o item: Criar Item → Finalizar → Salvar → Sim."""
    log.info("Criando item...")

    await safe_click(page, SELECTORS["criar_item_btn"])
    await page.wait_for_load_state("networkidle")

    await safe_click(page, SELECTORS["finalizar_btn"])
    await page.wait_for_load_state("networkidle")

    await safe_click(page, SELECTORS["salvar_btn"])
    await page.wait_for_load_state("networkidle")

    await safe_click(page, SELECTORS["sim_btn"])
    await page.wait_for_load_state("networkidle")

    log.info("Item criado com sucesso")


async def finalizar_e_remeter(page: Page) -> None:
    """Finaliza o item e remete para MODEC.

    Sequência: Finalizar → Atuar no Item → Remeter Modec → Sim
    """
    log.info("Finalizando e remetendo para MODEC...")

    # Finalizar
    await safe_click(page, SELECTORS["finalizar_btn"])
    await page.wait_for_load_state("networkidle")

    # Aguardar e clicar em "Atuar no Item"
    await page.wait_for_selector(SELECTORS["atuar_no_item_btn"], timeout=10_000)
    await safe_click(page, SELECTORS["atuar_no_item_btn"])
    await page.wait_for_load_state("networkidle")

    # Aguardar e clicar em "Remeter Modec"
    await page.wait_for_selector(SELECTORS["remeter_modec_btn"], timeout=10_000)
    await safe_click(page, SELECTORS["remeter_modec_btn"])
    await page.wait_for_load_state("networkidle")

    # Confirmar com "Sim"
    await safe_click(page, SELECTORS["sim_btn"])
    await page.wait_for_load_state("networkidle")

    log.info("Item remetido para MODEC")
