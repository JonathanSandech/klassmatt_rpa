"""Atributos técnicos — loop de atributos + preenchimento de popup."""

import asyncio

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click
from logger import log


def _attr_ctl_index(loop_index: int) -> str:
    """Converte índice do loop (1-based) para formato ASP.NET ctl{nn}.

    No ASP.NET DataGrid, os controles usam ctl02, ctl03, ctl04...
    O índice 1 do loop corresponde a ctl02 (primeira row de dados).
    """
    return f"{loop_index + 1:02d}"


async def fill_attributes(page: Page, attributes: list) -> None:
    """Preenche atributos técnicos (até 30).

    A tabela dgDadosTecnicos fica na página ITEM_Edita_DescricaoV3.aspx,
    acessada via Descrições → Editar Descrição. Se já estivermos nessa
    página (após change_pdm), não navega novamente.

    Para cada atributo:
    - Se vazio → para o loop (não há mais atributos)
    - Se "N/A" → marca checkbox N/A
    - Se tem valor → abre popup de árvore (Dt_EditaArvore.aspx) e seleciona
    """
    log.info("Preenchendo atributos técnicos...")

    # Garantir que estamos na página de edição de descrição (onde dgDadosTecnicos fica)
    if "ITEM_Edita_DescricaoV3" not in page.url:
        # Navegar via JS para evitar problemas de overlay
        await page.evaluate(
            """() => {
                const tabs = document.querySelectorAll('a');
                const tab = Array.from(tabs).find(a => a.innerText.includes('Descrições'));
                if (tab) tab.click();
            }"""
        )
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)

        # Clicar "Editar Descrição" via JS
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
        else:
            log.warning("Link 'Editar Descrição' não encontrado — pulando atributos")
            return

    # Verificar se a tabela de atributos existe
    has_grid = await page.locator("#dgDadosTecnicos").count() > 0
    if not has_grid:
        log.warning("Tabela dgDadosTecnicos não encontrada — pulando atributos")
        return

    for i, value in enumerate(attributes):
        if value is None or (isinstance(value, str) and value.strip() == ""):
            log.info(f"Atributo {i + 1}: vazio — encerrando loop de atributos")
            break

        value_str = str(value).strip()
        ctl_idx = _attr_ctl_index(i + 1)

        if value_str.upper() == "N/A":
            # Marcar checkbox N/A
            na_selector = SELECTORS["attr_na_checkbox_tpl"].format(idx=ctl_idx)
            na_el = page.locator(na_selector)
            if await na_el.count() == 0:
                log.debug(f"Atributo {i + 1}: checkbox N/A não encontrado (ctl{ctl_idx}) — fim dos atributos")
                break
            log.debug(f"Atributo {i + 1}: N/A")
            await safe_click(page, na_selector)
            await page.wait_for_timeout(1000)
        else:
            # Verificar se o botão de edição existe e está visível
            edit_selector = SELECTORS["attr_edit_btn_tpl"].format(idx=ctl_idx)
            edit_el = page.locator(edit_selector)
            if await edit_el.count() == 0:
                log.debug(f"Atributo {i + 1}: botão edição não encontrado (ctl{ctl_idx}) — fim dos atributos")
                break

            log.debug(f"Atributo {i + 1}: '{value_str}'")

            # Abrir popup da árvore via AbreJanTaxonomia() e preencher
            await _open_and_fill_tree_popup(page, ctl_idx, value_str)

            # Após popup fechar, verificar se ainda estamos na DescricaoV3
            # (selecionar na popup pode causar navegação na página pai)
            await page.wait_for_timeout(1000)
            if "ITEM_Edita_DescricaoV3" not in page.url:
                log.debug(f"Página mudou após popup: {page.url} — re-navegando para DescricaoV3")
                # Se saiu da DescricaoV3, precisamos navegar de volta
                if "ITEM_Resumo" in page.url or "SIN_Item_Resultante" in page.url:
                    atuar = page.locator(SELECTORS["atuar_no_item_btn"])
                    if await atuar.count() > 0:
                        await atuar.click()
                        await page.wait_for_load_state("networkidle")
                # Re-navegar para DescricaoV3 via JS
                if "ITEM_Edita_DescricaoV3" not in page.url:
                    await page.evaluate(
                        """() => {
                            const tabs = document.querySelectorAll('a');
                            const tab = Array.from(tabs).find(a => a.innerText.includes('Descrições'));
                            if (tab) tab.click();
                        }"""
                    )
                    await page.wait_for_load_state("networkidle")
                    await page.evaluate(
                        """() => {
                            const links = document.querySelectorAll('a');
                            const link = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
                            if (link) link.click();
                        }"""
                    )
                    await page.wait_for_load_state("networkidle")

    # Voltar para a página do item (ITEM_Edita.aspx) se estamos na DescricaoV3
    if "ITEM_Edita_DescricaoV3" in page.url:
        log.debug("Voltando da DescricaoV3 para a página do item...")
        voltar_btn = page.locator("#butSIN_Voltar")
        if await voltar_btn.count() > 0:
            await voltar_btn.click()
            await page.wait_for_load_state("networkidle", timeout=15_000)
            # butSIN_Voltar leva à SIN; precisamos clicar "Atuar no Item" de novo
            await page.wait_for_timeout(1000)
            atuar_btn = page.locator(SELECTORS["atuar_no_item_btn"])
            if await atuar_btn.count() > 0:
                await atuar_btn.click()
                await page.wait_for_load_state("networkidle", timeout=15_000)
                await page.wait_for_timeout(1000)
                log.debug("Voltou para ITEM_Edita.aspx via Atuar no Item")
            else:
                log.debug(f"Página atual após Voltar: {page.url}")

    log.info("Atributos técnicos preenchidos")


async def _open_and_fill_tree_popup(page: Page, ctl_idx: str, value: str) -> None:
    """Abre a popup de árvore (Dt_EditaArvore.aspx) e seleciona um valor.

    O botão btnAddEdit chama AbreJanTaxonomia() que abre window.open().
    O Playwright captura a nova janela via context.pages.

    A árvore tem estrutura hierárquica:
    - Nível 0: nome do dado técnico (ex: "NOME VALIDO", "APLICACAO")
    - Nível 1: letras do alfabeto ([0-9], A, B, C, ..., Z)
    - Nível 2+: valores reais (ex: "PORCA BORBOLETA" sob "P")

    Sequência:
    1. Abrir popup via AbreJanTaxonomia
    2. Clicar na letra correspondente à primeira letra do valor
    3. Esperar expansão e procurar o valor nos nós filhos
    4. Clicar no valor e em "Selecionar"
    """
    context = page.context
    name_suffix = f"dgDadosTecnicos$ctl{ctl_idx}$btnAddEdit"

    # Abrir popup via JS — tornar botão visível e chamar AbreJanTaxonomia
    await page.evaluate(
        """(nameSuffix) => {
            const btn = document.querySelector(`input[name$='${nameSuffix}']`);
            if (btn) {
                btn.style.display = 'inline';
                AbreJanTaxonomia(btn);
            }
        }""",
        name_suffix,
    )

    # Esperar a nova janela/aba aparecer
    popup_page = None
    for _ in range(20):  # até 10s
        await asyncio.sleep(0.5)
        for p in context.pages:
            if "Dt_EditaArvore" in p.url:
                popup_page = p
                break
        if popup_page:
            break

    if not popup_page:
        log.warning(f"Popup da árvore não abriu para ctl{ctl_idx} — pulando atributo")
        return

    try:
        await popup_page.wait_for_load_state("networkidle", timeout=15_000)

        first_letter = value[0].upper()

        # Passo 1: Clicar na letra do alfabeto para expandir a sub-árvore
        # As letras são links com __doPostBack que causam navegação/postback
        # O click causa full page reload — precisamos esperar a nova página carregar
        letter_found = await popup_page.evaluate(
            """(letter) => {
                const nodes = document.querySelectorAll('a[class*="nodeStyle"]');
                const letterNode = Array.from(nodes).find(a => a.innerText.trim() === letter);
                if (letterNode) {
                    letterNode.click();
                    return true;
                }
                return false;
            }""",
            first_letter,
        )
        if letter_found:
            log.debug(f"Expandindo letra '{first_letter}' na árvore")
            # __doPostBack causa full reload — esperar a página recarregar completamente
            try:
                await popup_page.wait_for_load_state("load", timeout=15_000)
            except Exception:
                pass
            await popup_page.wait_for_load_state("networkidle", timeout=15_000)
            await popup_page.wait_for_timeout(1500)
        else:
            log.debug(f"Letra '{first_letter}' não encontrada na árvore")

        # Passo 2: Procurar o valor nos nós expandidos
        # Usar evaluate pois pode haver milhares de nós (1900+)
        found = await popup_page.evaluate(
            """(value) => {
                const nodes = document.querySelectorAll('a.nodeStyle, a.nodeStyleSel');
                // Match exato
                let target = Array.from(nodes).find(a => a.innerText.trim() === value);
                // Match parcial se não achou exato
                if (!target) {
                    const upper = value.toUpperCase();
                    target = Array.from(nodes).find(a => a.innerText.trim().toUpperCase() === upper);
                }
                if (target) {
                    target.click();
                    return target.innerText.trim();
                }
                return null;
            }""",
            value,
        )

        if found:
            log.debug(f"Nó selecionado na árvore: '{found}'")
            await popup_page.wait_for_timeout(500)

            # Clicar em "Selecionar"
            sel_btn = popup_page.locator("#btnSelecionar")
            if await sel_btn.count() > 0:
                await sel_btn.click()
                # Esperar a popup fechar
                for _ in range(10):
                    await asyncio.sleep(0.5)
                    if popup_page.is_closed():
                        break
            log.debug(f"Popup preenchido: '{value}'")
        else:
            # Listar valores disponíveis para debug
            available = await popup_page.evaluate(
                """() => {
                    const nodes = document.querySelectorAll('a.nodeStyle, a.nodeStyleSel');
                    return Array.from(nodes).map(a => a.innerText.trim()).slice(0, 20);
                }"""
            )
            log.warning(f"Valor '{value}' não encontrado na árvore após expandir '{first_letter}'. Disponíveis: {available}")
            await popup_page.close()

    except Exception as e:
        log.warning(f"Erro ao preencher popup da árvore: {e}")
        try:
            if not popup_page.is_closed():
                await popup_page.close()
        except Exception:
            pass
