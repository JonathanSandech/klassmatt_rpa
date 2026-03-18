"""Aba Referências — empresa e part number."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, page_contains_text, wait_for_text
from logger import log


async def fill_reference(page: Page, empresa: str, part_number: str) -> bool:
    """Preenche referência (empresa + part number).

    Retorna True se ok, False se referência duplicada.
    """
    log.info(f"Preenchendo referência: {empresa} / {part_number}")

    # Navegar para aba Referências
    await safe_click(page, SELECTORS["tab_referencias"])
    await page.wait_for_load_state("networkidle")

    # Clicar no botão de adicionar referência (ID com I maiúsculo: Imagebutton22)
    add_btn = page.locator("input[type='image'][id$='Imagebutton22']")
    await add_btn.click()
    await page.wait_for_load_state("networkidle")

    # Preencher empresa (digitar para acionar autocomplete)
    await safe_fill(page, SELECTORS["ref_empresa_input"], str(empresa))
    await page.wait_for_timeout(2000)

    # Selecionar empresa na lista de sugestões (autocomplete)
    # Tentar match exato primeiro, depois parcial (primeiro resultado que contém o texto)
    empresa_option = page.locator(f"a:has-text('{empresa}')").first
    try:
        await empresa_option.click(timeout=5_000)
    except Exception:
        # Match parcial: clicar no primeiro resultado do autocomplete
        # O autocomplete mostra links <a> com class "ac_results" ou similar
        log.debug(f"Empresa '{empresa}' não encontrada exata, tentando primeiro resultado do autocomplete")
        first_option = page.locator(".ac_results a, .ui-autocomplete a, ul[id*='autocomplete'] a").first
        try:
            await first_option.click(timeout=5_000)
        except Exception:
            # Último fallback: clicar no primeiro <a> visível que apareceu após digitar
            fallback = page.locator(f"a:text-matches('{empresa.split()[0]}', 'i')").first
            await fallback.click(timeout=10_000)

    # Preencher Part Number
    await safe_fill(page, SELECTORS["ref_partnumber_input"], str(part_number))

    # Salvar
    salvar_btn = page.locator(SELECTORS["ref_salvar_btn"]).first
    await salvar_btn.click()
    await page.wait_for_load_state("networkidle")

    # Verificar duplicidade
    await page.wait_for_timeout(1000)
    is_duplicate = await page_contains_text(page, SELECTORS["ref_duplicate_text"])

    if is_duplicate:
        log.warning(f"Referência duplicada detectada: {empresa} / {part_number}")
        return False

    log.info("Referência salva com sucesso")
    return True
