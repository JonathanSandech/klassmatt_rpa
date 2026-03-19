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

    # Verificar se PDM já está definido de verdade.
    # dgDadosTecnicos SEMPRE existe (com TEXTO LONGO/TEXTO CURTO genéricos).
    # Só pular se Nome Válido NÃO for "(NÃO-PADRONIZADO)" — sinal de que um PDM real foi aplicado.
    pdm_is_set = await page.evaluate(
        """() => {
            const body = document.body.innerText;
            // Se "NÃO-PADRONIZADO" está na página, PDM não foi definido
            if (body.includes('NÃO-PADRONIZADO')) return false;
            // Verificar se existem dados técnicos reais (não só TEXTO LONGO/CURTO)
            const grid = document.querySelector('#dgDadosTecnicos');
            if (!grid) return false;
            const rows = grid.querySelectorAll('tr');
            for (const row of rows) {
                const text = row.innerText.trim();
                if (text && !text.includes('TEXTO LONGO') && !text.includes('TEXTO CURTO')
                    && !text.includes('Dados Técnicos') && !text.includes('NA') && text.length > 3) {
                    return true;  // Tem atributo real (ex: NOME VALIDO, APLICACAO)
                }
            }
            return false;
        }"""
    )
    if pdm_is_set:
        log.info("PDM já definido (atributos reais presentes) — pulando")
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
    # "Alterar Padrão" navega para Pesquisa_Item.aspx (página diferente)
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Preencher PDM no campo de busca (txtFiltro)
    pdm_input = page.locator("#txtFiltro")
    if await pdm_input.count() == 0:
        # Fallback: último input de texto visível
        pdm_input = page.locator("input[type='text']").last
    await pdm_input.fill(str(pdm))

    # Clicar em Pesquisar (não Enter — Enter pode causar form submit errado)
    await safe_click(page, "input[value='Pesquisar']")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Clicar em "PARTES E PECAS" via JS
    found_cat = await page.evaluate(
        f"""() => {{
            const links = document.querySelectorAll('a');
            const link = Array.from(links).find(a => a.innerText.includes('{PDM_CATEGORY}'));
            if (link) {{ link.click(); return true; }}
            return false;
        }}"""
    )
    if not found_cat:
        log.warning(f"Categoria '{PDM_CATEGORY}' não encontrada para PDM {pdm}")
        return
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Clicar em "Definir Padrão" via JS
    # Isso navega de volta para ITEM_Edita_DescricaoV3.aspx
    await page.evaluate(
        """() => {
            const btn = document.querySelector("input[value='Definir Padrão']");
            if (btn) btn.click();
        }"""
    )
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)

    # Garantir que voltamos para DescricaoV3 e a página está estável
    if "ITEM_Edita_DescricaoV3" not in page.url:
        log.debug(f"Após Definir Padrão, URL={page.url} — aguardando navegação")
        try:
            await page.wait_for_url("**/ITEM_Edita_DescricaoV3*", timeout=15_000)
        except Exception:
            log.warning(f"Não voltou para DescricaoV3 após Definir Padrão (URL={page.url})")
    await page.wait_for_load_state("domcontentloaded")
    await page.wait_for_timeout(2000)

    # NÃO clicar Finalizar aqui — Finalizar sem atributos preenchidos não salva o PDM.
    # O Finalizar será feito pelo fill_attributes() após preencher os atributos,
    # o que persiste PDM + atributos juntos.

    log.info(f"PDM alterado para: {pdm} / {PDM_CATEGORY}")
