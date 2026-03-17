"""Atributos técnicos — loop de atributos + preenchimento de popup."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click
from logger import log


def _attr_ctl_index(loop_index: int) -> str:
    """Converte índice do loop (1-based) para formato ASP.NET ctl{nn}.

    No ASP.NET DataGrid, os controles usam ctl02, ctl03, ctl04...
    O índice 1 do loop corresponde a ctl03 (ctl02 é o header).
    """
    return f"{loop_index + 2:02d}"


async def fill_attributes(page: Page, attributes: list) -> None:
    """Preenche atributos técnicos (até 30).

    Para cada atributo:
    - Se vazio → para o loop (não há mais atributos)
    - Se "N/A" → marca checkbox N/A
    - Se tem valor → clica no botão de edição e preenche via popup
    """
    log.info("Preenchendo atributos técnicos...")

    for i, value in enumerate(attributes):
        if value is None or (isinstance(value, str) and value.strip() == ""):
            log.info(f"Atributo {i + 1}: vazio — encerrando loop de atributos")
            break

        value_str = str(value).strip()
        ctl_idx = _attr_ctl_index(i + 1)

        if value_str.upper() == "N/A":
            # Marcar checkbox N/A
            na_selector = SELECTORS["attr_na_checkbox_tpl"].format(idx=ctl_idx)
            log.debug(f"Atributo {i + 1}: N/A")
            await safe_click(page, na_selector)
            await page.wait_for_timeout(1000)
        else:
            # Clicar no botão de edição do atributo
            edit_selector = SELECTORS["attr_edit_btn_tpl"].format(idx=ctl_idx)
            log.debug(f"Atributo {i + 1}: '{value_str}'")
            await safe_click(page, edit_selector)
            await page.wait_for_timeout(2000)

            # Preencher popup
            await _fill_attribute_popup(page, value_str)

    log.info("Atributos técnicos preenchidos")


async def _fill_attribute_popup(page: Page, value: str) -> None:
    """Preenche o popup de seleção de atributo.

    Diferente do PAD que usa DevTools + JS injection, o Playwright
    pode acessar os elementos do popup diretamente via evaluate().

    Sequência:
    1. Extrair primeira letra → maiúscula
    2. Clicar na letra no alfabeto do popup
    3. Clicar no valor na lista
    4. Clicar em "Selecionar"
    """
    first_letter = value[0].upper()

    # O popup pode estar em um iframe ou na mesma página.
    # Tentar primeiro com evaluate (como o PAD fazia via DevTools)
    try:
        # Clicar na letra do alfabeto
        await page.evaluate(
            """(letter) => {
                const el = Array.from(document.querySelectorAll('.txt-letra'))
                    .find(e => e.innerText.trim() === letter);
                if (el) el.closest('a').click();
            }""",
            first_letter,
        )
        await page.wait_for_timeout(2000)

        # Clicar no valor
        await page.evaluate(
            """(value) => {
                const el = Array.from(document.querySelectorAll('a.nodeStyle'))
                    .find(e => e.innerText.trim() === value);
                if (el) el.click();
            }""",
            value,
        )
        await page.wait_for_timeout(1000)

        # Clicar em "Selecionar"
        await page.evaluate(
            """() => {
                const btn = document.getElementById('btnSelecionar');
                if (btn) btn.click();
            }"""
        )
        await page.wait_for_timeout(1000)

    except Exception as e:
        # Fallback: tentar com seletores Playwright diretos
        log.warning(f"Popup evaluate falhou, tentando seletores diretos: {e}")

        letter_sel = SELECTORS["popup_letter_tpl"].format(letter=first_letter)
        await safe_click(page, letter_sel)
        await page.wait_for_timeout(2000)

        value_sel = SELECTORS["popup_value_tpl"].format(value=value)
        await safe_click(page, value_sel)
        await page.wait_for_timeout(1000)

        await safe_click(page, SELECTORS["popup_select_btn"])
        await page.wait_for_timeout(1000)

    log.debug(f"Popup preenchido: '{value}'")
