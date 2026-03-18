"""Aba Referências — empresa e part number."""

import re

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, page_contains_text, wait_for_text
from logger import log


async def _has_existing_reference(page: Page) -> bool:
    """Verifica se já existe pelo menos uma referência pelo label da aba.

    A aba mostra 'Referências (N)' — se N > 0, já existe.
    """
    tab_el = page.locator("a:has-text('Referências')").first
    try:
        tab_text = await tab_el.inner_text(timeout=5_000)
        match = re.search(r"\((\d+)\)", tab_text)
        if match and int(match.group(1)) > 0:
            return True
    except Exception:
        pass
    return False


async def _select_autocomplete(page: Page, empresa: str) -> bool:
    """Tenta selecionar a empresa no autocomplete. Retorna True se selecionou."""

    # 1. Match exato via texto
    try:
        await page.locator(f"a:has-text('{empresa}')").first.click(timeout=3_000)
        return True
    except Exception:
        pass

    # 2. Qualquer item visível no autocomplete (jQuery UI, ac_results)
    for ac_sel in [
        ".ac_results li:first-child a",
        ".ac_results li:first-child",
        ".ui-autocomplete li:first-child a",
        ".ui-autocomplete li:first-child",
    ]:
        try:
            el = page.locator(ac_sel).first
            if await el.count() > 0 and await el.is_visible():
                await el.click(timeout=3_000)
                return True
        except Exception:
            continue

    # 3. JS fallback — qualquer lista de autocomplete visível
    try:
        clicked = await page.evaluate(
            """() => {
                const lists = document.querySelectorAll('.ac_results, .ui-autocomplete, [id*="autocomplete"]');
                for (const list of lists) {
                    if (list.offsetParent !== null) {
                        const items = list.querySelectorAll('li, a');
                        if (items.length > 0) { items[0].click(); return true; }
                    }
                }
                return false;
            }"""
        )
        if clicked:
            return True
    except Exception:
        pass

    # 4. Match parcial com primeira palavra
    first_word = empresa.split()[0] if empresa.split() else empresa
    try:
        await page.locator(f"a:text-matches('{first_word}', 'i')").first.click(timeout=5_000)
        return True
    except Exception:
        pass

    return False


async def fill_reference(page: Page, empresa: str, part_number: str) -> bool:
    """Preenche referência (empresa + part number).

    Retorna True se ok, False se referência duplicada.
    """
    log.info(f"Preenchendo referência: {empresa} / {part_number}")

    # Navegar para aba Referências
    await safe_click(page, SELECTORS["tab_referencias"])
    await page.wait_for_load_state("networkidle")

    # Verificar se já existe referência — se sim, pular (idempotente)
    if await _has_existing_reference(page):
        log.info("Referência já existe — pulando")
        return True

    # Clicar no botão ADD para nova referência (iButAddRef, não Imagebutton22 que é EDIT)
    add_btn = page.locator("#iButAddRef")
    if await add_btn.count() > 0 and await add_btn.is_visible():
        await add_btn.click()
        await page.wait_for_load_state("networkidle")
    else:
        # Fallback: usar Imagebutton22 se iButAddRef não existe (item recém-criado com row vazia)
        edit_btn = page.locator("input[type='image'][id$='Imagebutton22']")
        await edit_btn.click()
        await page.wait_for_load_state("networkidle")

    # Preencher empresa (digitar para acionar autocomplete)
    await safe_fill(page, SELECTORS["ref_empresa_input"], str(empresa))
    await page.wait_for_timeout(2000)

    # Selecionar empresa no autocomplete
    autocomplete_ok = await _select_autocomplete(page, empresa)
    if not autocomplete_ok:
        log.warning(f"Autocomplete empresa '{empresa}' não encontrado — continuando sem seleção")

    # Preencher Part Number
    await safe_fill(page, SELECTORS["ref_partnumber_input"], str(part_number))

    # Salvar — o confirm dialog "Fabricante não existe, deseja cadastrá-lo?" é aceito
    # automaticamente pelo handler global.
    salvar_btn = page.locator(SELECTORS["ref_salvar_btn"]).first
    await salvar_btn.click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Verificar duplicidade
    is_duplicate = await page_contains_text(page, SELECTORS["ref_duplicate_text"])
    if is_duplicate:
        log.warning(f"Referência duplicada detectada: {empresa} / {part_number}")
        return False

    # Após salvar com fabricante novo, o form pode continuar aberto (dirty state).
    # Tentar salvar novamente se #btnSalvar ainda estiver visível.
    for retry in range(2):
        try:
            salvar_still = page.locator(SELECTORS["ref_salvar_btn"]).first
            if await salvar_still.is_visible(timeout=1_500):
                log.debug(f"Form referência ainda aberto (retry {retry + 1}) — salvando novamente")
                await salvar_still.click()
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)
            else:
                break
        except Exception:
            break

    log.info("Referência salva com sucesso")
    return True
