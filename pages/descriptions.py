"""Aba Descrições — validação SAP (Exibe D2 / 40 chars) + alteração PDM."""

import re

from playwright.async_api import Page

from config import SELECTORS, PDM_CATEGORY
from browser import safe_click, safe_fill
from logger import log


async def validate_sap_description(page: Page) -> None:
    """Verifica tamanho da descrição SAP e desmarca 'Exibe D2' se > 40 chars.

    Lê o texto da descrição SAP que contém o padrão "(tam: XX/40)",
    extrai o número antes da / e se > 40, vai em Referências para
    desmarcar o checkbox Exibe D2.
    """
    log.info("Validando descrição SAP...")

    # Navegar para aba Descrições
    await safe_click(page, SELECTORS["tab_descricoes"])
    await page.wait_for_load_state("networkidle")

    # Ler texto que contém o tamanho (ex: "NUT ... HUGHES (tam: 55/40)")
    # O PAD usa regex \d+(?=\/) para extrair o número antes da barra
    page_text = await page.inner_text("body")
    match = re.search(r"tam:\s*(\d+)/", page_text)

    if not match:
        log.warning("Não encontrou padrão de tamanho SAP — salvando normalmente")
        salvar_btn = page.locator(SELECTORS["salvar_btn"])
        if await salvar_btn.count() > 0:
            await salvar_btn.nth(1).click()  # "Salvar 2" no PAD
            await page.wait_for_load_state("networkidle")
            await page.keyboard.press("Enter")
        return

    tamanho = int(match.group(1))
    log.info(f"Tamanho da descrição SAP: {tamanho}/40")

    if tamanho > 40:
        log.info("Tamanho > 40 — desmarcando 'Exibe D2'")

        # Ir para aba Referências
        await safe_click(page, SELECTORS["tab_referencias"])
        await page.wait_for_load_state("networkidle")

        # Editar referência (botão de edição)
        edit_btn = page.locator("input[type='image'][id$='imagebutton22']")
        await edit_btn.click()
        await page.wait_for_load_state("networkidle")

        # Desmarcar checkbox Exibe D2
        checkbox = page.locator(SELECTORS["ref_exibe_d2_checkbox"])
        if await checkbox.is_checked():
            await checkbox.uncheck()

        # Salvar
        salvar_btn = page.locator(SELECTORS["ref_salvar_btn"]).first
        await salvar_btn.click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2000)

        # Segundo salvar (confirmação)
        salvar2 = page.locator(SELECTORS["salvar_btn"])
        if await salvar2.count() > 1:
            await salvar2.nth(1).click()
            await page.wait_for_load_state("networkidle")

        await page.keyboard.press("Enter")
    else:
        # Apenas salvar
        salvar2 = page.locator(SELECTORS["salvar_btn"])
        if await salvar2.count() > 1:
            await salvar2.nth(1).click()
            await page.wait_for_load_state("networkidle")
        await page.keyboard.press("Enter")

    log.info("Validação SAP concluída")


async def change_pdm(page: Page, pdm: str) -> None:
    """Altera o padrão PDM na aba Descrições.

    Sequência: Descrições → Editar Descrição → Alterar Padrão →
    digitar PDM → Enter → clicar 'PARTES E PECAS' → Definir Padrão
    """
    log.info(f"Alterando PDM para: {pdm}")

    # Navegar para aba Descrições
    await safe_click(page, SELECTORS["tab_descricoes"])
    await page.wait_for_load_state("networkidle")

    # Clicar em "Editar Descrição"
    await safe_click(page, SELECTORS["editar_descricao_link"])
    await page.wait_for_load_state("networkidle")

    # Aguardar botão "Alterar Padrão"
    await page.wait_for_selector(SELECTORS["alterar_padrao_btn"], timeout=10_000)
    await safe_click(page, SELECTORS["alterar_padrao_btn"])
    await page.wait_for_timeout(1000)

    # Preencher PDM no campo de busca (campo de texto que aparece)
    # O PAD usava UIAutomation para isso — aqui tentamos o input que aparece
    pdm_input = page.locator("input[type='text']").last
    await pdm_input.fill(str(pdm))
    await page.wait_for_timeout(2000)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(2000)

    # Clicar em "PARTES E PECAS"
    await safe_click(page, SELECTORS["partes_pecas_link"])
    await page.wait_for_timeout(2000)

    # Clicar em "Definir Padrão"
    await safe_click(page, SELECTORS["definir_padrao_btn"])
    await page.wait_for_load_state("networkidle")

    log.info(f"PDM alterado para: {pdm} / {PDM_CATEGORY}")
