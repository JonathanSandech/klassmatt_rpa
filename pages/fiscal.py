"""Aba Fiscal — preenchimento do NCM."""

import browser as _browser
from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill
from logger import log


async def fill_ncm(page: Page, ncm: str) -> None:
    """Navega para aba Fiscal e preenche o campo NCM/TIPI.

    Após preencher, dispara Tab para provocar a validação do ASP.NET.
    Detecta rejeição via browser.last_dialog_message (o handler global
    grava a última mensagem). Se rejeitado, limpa o campo para evitar
    cascata de alerts nas próximas trocas de aba.
    """
    log.info(f"Preenchendo NCM: {ncm}")

    await safe_click(page, SELECTORS["tab_fiscal"])
    await page.wait_for_load_state("networkidle")

    # Verificar se o campo já está preenchido e readonly (item parcialmente processado)
    ncm_el = page.locator(SELECTORS["ncm_input"])
    is_readonly = await ncm_el.get_attribute("readonly")
    current_value = await ncm_el.input_value()
    if is_readonly and current_value:
        log.info(f"NCM já preenchido ({current_value}) e readonly — pulando")
        return

    # Limpar flag de dialog antes de preencher
    _browser.last_dialog_message = ""

    await safe_fill(page, SELECTORS["ncm_input"], str(ncm))

    # Disparar validação saindo do campo (Tab) e aguardar postback
    await page.keyboard.press("Tab")
    await page.wait_for_timeout(3000)

    # Checar se alert de NCM invalido foi capturado pelo dialog handler
    last_msg = _browser.last_dialog_message.lower()
    ncm_rejected = "ncm" in last_msg or "inválido" in last_msg or "inativo" in last_msg

    if not ncm_rejected:
        # Fallback: se o campo ficou vazio, foi rejeitado
        ncm_value = await page.input_value(SELECTORS["ncm_input"])
        if not ncm_value or ncm_value.strip() == "":
            ncm_rejected = True

    if ncm_rejected:
        log.warning(f"NCM {ncm} rejeitado pelo sistema -- limpando campo")
        await page.fill(SELECTORS["ncm_input"], "")
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(2000)
        # Limpar qualquer alert residual
        _browser.last_dialog_message = ""
    else:
        log.info(f"NCM {ncm} preenchido")
