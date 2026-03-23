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
    await page.wait_for_timeout(500)
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
        log.info("Tamanho > 40 — verificando Exibe D2")

        await _click_tab(page, "Referências")

        # Verificar se existe referência e botão de edição antes de abrir o form
        has_edit_btn = await page.evaluate(
            """() => !!document.querySelector("[id$='Imagebutton22']")"""
        )
        if not has_edit_btn:
            log.debug("Sem referência para editar Exibe D2 — pulando")
            log.info("Validação SAP concluída")
            return

        # Editar referência existente via JS
        await page.evaluate(
            """() => {
                const btn = document.querySelector("[id$='Imagebutton22']");
                if (btn) btn.click();
            }"""
        )
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(500)
        await hide_overlays(page)

        # Desmarcar checkbox Exibe D2
        checkbox = page.locator(SELECTORS["ref_exibe_d2_checkbox"])
        d2_changed = False
        try:
            if await checkbox.count() > 0 and await checkbox.is_checked():
                await checkbox.uncheck()
                d2_changed = True
                log.debug("Exibe D2 desmarcado")
            else:
                log.debug("Exibe D2 já desmarcado ou não encontrado")
        except Exception:
            log.debug("Checkbox Exibe D2 não acessível")

        if d2_changed:
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
            for btn_selector in [
                "input[value='Continuar']",
                "input[value='continuar']",
                "a:has-text('Continuar')",
            ]:
                btn = page.locator(btn_selector)
                try:
                    if await btn.count() > 0 and await btn.is_visible():
                        log.debug("Aviso 'Referência igual' — clicando 'Continuar'")
                        await btn.click()
                        try:
                            await page.wait_for_load_state("networkidle", timeout=10_000)
                        except Exception:
                            pass
                        await page.wait_for_timeout(500)
                        break
                except Exception:
                    continue
        else:
            # Não mudou nada — cancelar edição da referência
            # Override confirm para aceitar "alterações não salvas" automaticamente
            await page.evaluate(
                """() => {
                    window.confirm = () => true;
                    window.alert = () => {};
                    // Clicar no botão Adicionar (iButAddRef) força sair do modo edição
                    // Ou navegar via __doPostBack para Dados Básicos
                    const tabs = document.querySelectorAll('a');
                    const tab = Array.from(tabs).find(a => a.innerText.includes('Dados Básicos'));
                    if (tab) tab.click();
                }"""
            )
            try:
                await page.wait_for_load_state("networkidle", timeout=10_000)
            except Exception:
                pass
            await page.wait_for_timeout(500)

        # Garantir que voltamos à página de edição do item (não ficamos em página de aviso)
        if "ITEM_Edita" not in page.url:
            log.debug(f"Após Exibe D2, URL inesperada: {page.url} — tentando voltar")
            voltar_btn = page.locator("input[value='Voltar']")
            if await voltar_btn.count() > 0:
                await voltar_btn.click()
                try:
                    await page.wait_for_load_state("networkidle", timeout=10_000)
                except Exception:
                    pass

    log.info("Validação SAP concluída")


async def change_pdm(page: Page, pdm: str) -> bool:
    """Altera o padrão PDM na aba Descrições.

    Sequência: Descrições → Editar Descrição → Alterar Padrão →
    digitar PDM → Enter → clicar 'PARTES E PECAS' → Definir Padrão

    Retorna True se PDM foi alterado ou já estava correto, False se falhou.
    """
    log.info(f"Alterando PDM para: {pdm}")

    await _click_tab(page, "Descrições")
    await page.wait_for_timeout(1000)

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
            await page.wait_for_timeout(500)
            await hide_overlays(page)
        else:
            log.warning("Link 'Editar Descrição' não encontrado — PDM pode já estar definido")
            return False

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
        return True

    # Aguardar botão "Alterar Padrão"
    alterar_btn = page.locator(SELECTORS["alterar_padrao_btn"])
    try:
        await alterar_btn.wait_for(state="visible", timeout=10_000)
    except Exception:
        log.warning("Botão 'Alterar Padrão' não encontrado — PDM pode já estar definido")
        return False

    try:
        await page.evaluate(
            """() => {
                const btn = document.querySelector("input[value='Alterar Padrão']");
                if (btn) btn.click();
            }"""
        )
    except Exception:
        pass  # Esperado: navegação destrói contexto, mas o clique já foi disparado
    # "Alterar Padrão" navega para Pesquisa_Item.aspx (página diferente)
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Preencher PDM no campo de busca (txtFiltro)
    pdm_input = page.locator("#txtFiltro")
    if await pdm_input.count() == 0:
        # Fallback: último input de texto visível
        pdm_input = page.locator("input[type='text']").last
    await pdm_input.fill(str(pdm))

    # Clicar em Pesquisar (não Enter — Enter pode causar form submit errado)
    await safe_click(page, "input[value='Pesquisar']")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Clicar na categoria do resultado da pesquisa (nome válido do PDM)
    # Cada PDM tem sua categoria (17100=PARTES E PECAS, 18101=KIT, etc)
    # Clicar no primeiro link da tabela de resultados que é o nome válido
    found_cat = await page.evaluate(
        f"""() => {{
            // Primeiro: tentar encontrar link com o número do PDM (ex: "17100")
            const links = document.querySelectorAll('a');
            const pdmLink = Array.from(links).find(a => a.innerText.trim() === '{pdm}');
            if (pdmLink) {{ pdmLink.click(); return pdmLink.innerText.trim(); }}
            // Fallback: clicar no segundo link da primeira row de dados (nome válido)
            const table = document.querySelector('table[id*="dgPadroes"]');
            if (table) {{
                const rows = table.querySelectorAll('tr');
                for (const row of rows) {{
                    const rowLinks = row.querySelectorAll('a');
                    if (rowLinks.length >= 2) {{
                        rowLinks[1].click();
                        return rowLinks[1].innerText.trim();
                    }}
                }}
            }}
            return null;
        }}"""
    )
    if not found_cat:
        log.warning(f"Categoria não encontrada para PDM {pdm}")
        return False
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Clicar em "Definir Padrão" via JS
    # Isso navega de volta para ITEM_Edita_DescricaoV3.aspx
    # O clique causa navegação que destrói o contexto JS antes do evaluate retornar
    try:
        await page.evaluate(
            """() => {
                const btn = document.querySelector("input[value='Definir Padrão']");
                if (btn) btn.click();
            }"""
        )
    except Exception:
        pass  # Esperado: navegação destrói contexto, mas o clique já foi disparado
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1500)

    # Garantir que voltamos para DescricaoV3 e a página está estável
    if "ITEM_Edita_DescricaoV3" not in page.url:
        log.debug(f"Após Definir Padrão, URL={page.url} — aguardando navegação")
        try:
            await page.wait_for_url("**/ITEM_Edita_DescricaoV3*", timeout=15_000)
        except Exception:
            log.warning(f"Não voltou para DescricaoV3 após Definir Padrão (URL={page.url})")
    await page.wait_for_timeout(1000)

    # NÃO clicar Finalizar aqui — Finalizar sem atributos preenchidos não salva o PDM.
    # O Finalizar será feito pelo fill_attributes() após preencher os atributos,
    # o que persiste PDM + atributos juntos.

    log.info(f"PDM alterado para: {pdm} / {found_cat}")
    return True
