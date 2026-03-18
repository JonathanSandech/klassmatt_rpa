"""Aba Descrições — validação SAP (Exibe D2 / 40 chars) + alteração PDM."""

import re

from playwright.async_api import Page

from config import SELECTORS, PDM_CATEGORY
from browser import safe_click, safe_fill, hide_overlays
from logger import log


async def _click_tab(page: Page, tab_name: str) -> None:
    """Clica em uma aba via JS para evitar problemas de overlay."""
    await page.evaluate(
        f"""() => {{
            const tabs = document.querySelectorAll('a');
            const tab = Array.from(tabs).find(a => a.innerText.includes('{tab_name}'));
            if (tab) tab.click();
        }}"""
    )
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)
    await hide_overlays(page)


async def validate_sap_description(page: Page) -> None:
    """Verifica tamanho da descrição SAP e desmarca 'Exibe D2' se > 40 chars."""
    log.info("Validando descrição SAP...")

    await _click_tab(page, "Descrições")

    # Ler descrição SAP (D2)
    try:
        d2_text = await page.inner_text("#txtD2")
    except Exception:
        d2_text = await page.inner_text("body")
    match = re.search(r"tam:\s*(\d+)/", d2_text)

    if not match:
        log.warning("Não encontrou padrão de tamanho SAP — continuando")
        return

    tamanho = int(match.group(1))
    log.info(f"Tamanho da descrição SAP: {tamanho}/40")

    if tamanho > 40:
        log.info("Tamanho > 40 — desmarcando 'Exibe D2'")

        await _click_tab(page, "Referências")

        # Editar referência existente via JS
        await page.evaluate(
            """() => {
                const btn = document.querySelector("[id$='Imagebutton22']");
                if (btn) btn.click();
            }"""
        )
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
        await hide_overlays(page)

        # Desmarcar checkbox Exibe D2
        checkbox = page.locator(SELECTORS["ref_exibe_d2_checkbox"])
        try:
            if await checkbox.count() > 0 and await checkbox.is_checked():
                await checkbox.uncheck()
                log.debug("Exibe D2 desmarcado")
            else:
                log.debug("Exibe D2 já desmarcado ou não encontrado")
        except Exception:
            log.debug("Checkbox Exibe D2 não acessível")

        # Salvar referência via JS
        await page.evaluate(
            """() => {
                const btn = document.querySelector('#btnSalvar');
                if (btn) btn.click();
            }"""
        )
        # Timeout curto — o save pode redirecionar para página de aviso
        try:
            await page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass
        await page.wait_for_timeout(2000)

        # Verificar se apareceu aviso "Referência igual" com botão Continuar/Voltar
        # Tentar múltiplos seletores pois o layout pode variar
        for btn_selector in [
            "input[value='Continuar']",
            "input[value='continuar']",
            "a:has-text('Continuar')",
            "input[value='Voltar']",
        ]:
            btn = page.locator(btn_selector)
            try:
                if await btn.count() > 0 and await btn.is_visible():
                    btn_val = btn_selector.split("'")[1] if "'" in btn_selector else btn_selector
                    log.debug(f"Aviso detectado — clicando '{btn_val}'")
                    if "Continuar" in btn_selector:
                        await btn.click()
                    else:
                        # Se só tem Voltar, clicar para voltar à página do item
                        await btn.click()
                    try:
                        await page.wait_for_load_state("networkidle", timeout=10_000)
                    except Exception:
                        pass
                    await page.wait_for_timeout(1000)
                    break
            except Exception:
                continue

    log.info("Validação SAP concluída")


async def change_pdm(page: Page, pdm: str) -> None:
    """Altera o padrão PDM na aba Descrições.

    Sequência: Descrições → Editar Descrição → Alterar Padrão →
    digitar PDM → Enter → clicar 'PARTES E PECAS' → Definir Padrão
    """
    log.info(f"Alterando PDM para: {pdm}")

    await _click_tab(page, "Descrições")

    # Verificar se PDM já está definido (idempotente)
    pdm_already_set = await page.evaluate(
        """() => {
            const links = document.querySelectorAll('a');
            const editLink = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
            // Se não há link "Editar Descrição" mas há conteúdo de descrição, PDM pode já estar set
            const padrao = document.querySelector('#txtPadrao');
            if (padrao && padrao.value && padrao.value !== '1') return true;
            return false;
        }"""
    )

    # Clicar em "Editar Descrição" via JS
    if "ITEM_Edita_DescricaoV3" not in page.url:
        found = await page.evaluate(
            """() => {
                const links = document.querySelectorAll('a');
                const link = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
                if (link) { link.click(); return true; }
                return false;
            }"""
        )
        if found:
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
            await hide_overlays(page)
        else:
            log.warning("Link 'Editar Descrição' não encontrado — PDM pode já estar definido")
            return

    # Verificar se já existe dgDadosTecnicos (PDM já definido com atributos)
    has_grid = await page.locator("#dgDadosTecnicos").count() > 0
    if has_grid:
        log.info(f"PDM já definido (dgDadosTecnicos presente) — pulando")
        return

    # Aguardar botão "Alterar Padrão"
    alterar_btn = page.locator(SELECTORS["alterar_padrao_btn"])
    try:
        await alterar_btn.wait_for(state="visible", timeout=10_000)
    except Exception:
        log.warning("Botão 'Alterar Padrão' não encontrado — PDM pode já estar definido")
        return

    await page.evaluate(
        """() => {
            const btn = document.querySelector("input[value='Alterar Padrão']");
            if (btn) btn.click();
        }"""
    )
    await page.wait_for_timeout(2000)

    # Preencher PDM no campo de busca
    pdm_input = page.locator("input[type='text']").last
    await pdm_input.fill(str(pdm))
    await page.wait_for_timeout(2000)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(2000)

    # Clicar em "PARTES E PECAS" via JS
    await page.evaluate(
        f"""() => {{
            const links = document.querySelectorAll('a');
            const link = Array.from(links).find(a => a.innerText.includes('{PDM_CATEGORY}'));
            if (link) link.click();
        }}"""
    )
    await page.wait_for_timeout(2000)

    # Clicar em "Definir Padrão" via JS
    await page.evaluate(
        """() => {
            const btn = document.querySelector("input[value='Definir Padrão']");
            if (btn) btn.click();
        }"""
    )
    await page.wait_for_load_state("networkidle")

    log.info(f"PDM alterado para: {pdm} / {PDM_CATEGORY}")
