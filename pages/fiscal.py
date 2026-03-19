"""Aba Fiscal — preenchimento do NCM."""

import re

import browser as _browser
from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill
from logger import log


def _format_ncm(ncm: str) -> str:
    """Formata NCM de 8 dígitos para o padrão XXXX.XX.XX esperado pelo Klassmatt.

    Ex: '73181500' → '7318.15.00', '84841000' → '8484.10.00'
    Se já tiver pontos ou não tiver 8 dígitos, retorna como está.
    """
    digits = re.sub(r"\D", "", ncm)
    if len(digits) == 8:
        return f"{digits[:4]}.{digits[4:6]}.{digits[6:8]}"
    return ncm


async def fill_ncm(page: Page, ncm: str) -> None:
    """Navega para aba Fiscal e preenche o campo NCM/TIPI.

    Após preencher, dispara Tab para provocar a validação do ASP.NET.
    Detecta rejeição via browser.last_dialog_message (o handler global
    grava a última mensagem). Se rejeitado, limpa o campo para evitar
    cascata de alerts nas próximas trocas de aba.
    """
    ncm_formatted = _format_ncm(str(ncm))
    log.info(f"Preenchendo NCM: {ncm} → {ncm_formatted}")

    await safe_click(page, SELECTORS["tab_fiscal"])
    await page.wait_for_load_state("networkidle")

    # Verificar se o campo já está preenchido e readonly (item parcialmente processado)
    ncm_el = page.locator(SELECTORS["ncm_input"])
    is_editable = await ncm_el.is_editable()
    current_value = await ncm_el.input_value()
    if not is_editable:
        if current_value.strip() == ncm_formatted:
            log.info(f"NCM campo não editável, valor correto ({current_value}) — pulando")
        else:
            log.warning(
                f"NCM campo não editável com valor DIFERENTE do esperado: "
                f"'{current_value}' (esperado '{ncm_formatted}') — não é possível corrigir"
            )
        return
    if current_value and current_value.strip() == ncm_formatted:
        log.info(f"NCM já preenchido corretamente ({current_value}) — pulando")
        return
    if current_value and current_value.strip():
        log.warning(
            f"NCM existente '{current_value}' difere do esperado '{ncm_formatted}' — substituindo"
        )

    # Limpar flag de dialog antes de preencher
    _browser.last_dialog_message = ""

    await safe_fill(page, SELECTORS["ncm_input"], ncm_formatted)

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
        _browser.last_dialog_message = ""
    else:
        # Salvar para persistir o NCM (sem Salvar, valor é perdido ao trocar de aba)
        await safe_click(page, SELECTORS["salvar_btn"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
        log.info(f"NCM {ncm} preenchido e salvo")
