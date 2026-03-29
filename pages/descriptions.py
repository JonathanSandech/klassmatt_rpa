"""Aba Descrições — validação SAP (Exibe D2 / 40 chars) + alteração PDM."""

import re

from playwright.async_api import Page

from config import SELECTORS, PDM_CATEGORY
from browser import safe_click, safe_fill, hide_overlays
from logger import log


async def _click_tab(page: Page, tab_name: str) -> None:
    """Clica em uma aba via JS com override de alert/confirm."""
    await page.evaluate(
        f"""() => {{
            window.confirm = () => true;
            window.alert = () => {{}};
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
                    window.alert = () => {};
                    window.confirm = () => true;
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
            # Não mudou nada — cancelar edição da referência para limpar dirty state
            await page.evaluate(
                """() => {
                    window.confirm = () => true;
                    window.alert = () => {};
                    // Cancelar o form de edição da referência (limpa dirty state)
                    const cancelar = document.querySelector('#btnCancelar');
                    if (cancelar) { cancelar.click(); return; }
                    // Fallback: navegar para Dados Básicos
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

    # Se já estamos em DescricaoV3, pular direto para a verificação de PDM
    if "ITEM_Edita_DescricaoV3" not in page.url:
        # Garantir que estamos em ITEM_Edita.aspx (não em SIN_Item_Resultante)
        if "ITEM_Edita.aspx" not in page.url:
            log.warning(f"change_pdm: URL inesperada {page.url} — tentando navegar para ITEM_Edita")
            # Tentar clicar "Atuar no Item" se estamos no resumo da SIN
            atuar = await page.evaluate("""() => {
                const btn = document.querySelector("input[value='Atuar no Item']");
                if (btn && !btn.disabled) { btn.click(); return true; }
                return false;
            }""")
            if atuar:
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1500)
                await hide_overlays(page)

        if "ITEM_Edita.aspx" in page.url:
            # Navegar para aba Descrições via __doPostBack (mais confiável que tab.click)
            await page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
            }""")
            await hide_overlays(page)
            await page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
                // Encontrar o __doPostBack correto para a aba Descrições
                const tabs = document.querySelectorAll('a');
                const tab = Array.from(tabs).find(a => {
                    const text = a.innerText.trim();
                    return text === 'Descrições' && a.href && a.href.includes('__doPostBack');
                });
                if (tab) tab.click();
            }""")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
            await hide_overlays(page)

            # Verificar se a aba Descrições realmente carregou (link "Editar Descrição" presente)
            has_edit_link = await page.evaluate("""() => {
                const links = document.querySelectorAll('a');
                return !!Array.from(links).find(a => a.innerText.includes('Editar Descri'));
            }""")
            if not has_edit_link:
                log.warning("Aba Descrições não carregou (link 'Editar Descrição' ausente) — retry via __doPostBack")
                # Retry: usar __doPostBack direto para a aba Descrições (ctl01)
                await page.evaluate("""() => {
                    window.confirm = () => true;
                    window.alert = () => {};
                    __doPostBack('ctl00$Body$dlTab$ctl01$lbutMenu', '');
                }""")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)
                await hide_overlays(page)

            # Clicar em "Editar Descrição" via __doPostBack direto
            found = await page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
                // Tentar via link com texto
                const links = document.querySelectorAll('a');
                const link = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
                if (link) { link.click(); return 'link'; }
                // Fallback: __doPostBack direto
                try {
                    __doPostBack('ctl00$Body$tabDescricoes$lbutAlterarDescr', '');
                    return 'postback';
                } catch(e) { return null; }
            }""")
            if found:
                log.debug(f"Editar Descrição clicado via: {found}")
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)
                await hide_overlays(page)
            else:
                log.warning("Link 'Editar Descrição' não encontrado e __doPostBack falhou")
                return False

        # Verificar se chegamos em DescricaoV3
        if "ITEM_Edita_DescricaoV3" not in page.url:
            log.warning(f"Não navegou para DescricaoV3 (URL={page.url})")
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

    # Salvar o IdItem ANTES de navegar para Pesquisa_Item.aspx
    # (necessário porque Definir Padrão volta com IdItem=0, causando NullReferenceException)
    id_item = await page.evaluate(
        """() => {
            const url = window.location.href;
            const m = url.match(/IdItem=(\\d+)/i);
            if (m && m[1] !== '0') return m[1];
            // Fallback: buscar no body text
            const body = document.body.innerText;
            const m2 = body.match(/IdItem:\\s*(\\d+)/);
            return m2 ? m2[1] : null;
        }"""
    )
    log.debug(f"IdItem capturado antes de Alterar Padrão: {id_item}")

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

    # FIX: "Definir Padrão" retorna com IdItem=0 na URL, o que causa
    # NullReferenceException ao clicar Finalizar. Solução: voltar para
    # ITEM_Edita.aspx e re-entrar na DescricaoV3 pelo fluxo normal,
    # que gera a URL com o IdItem correto.
    if "IdItem=0" in page.url or "ITEM_AlterarPD=1" in page.url:
        log.info("URL com IdItem=0 detectada — re-navegando via fluxo normal")

        # Voltar à SIN
        voltar_btn = page.locator("#butSIN_Voltar")
        if await voltar_btn.count() > 0:
            await voltar_btn.click()
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)

        # Clicar "Atuar no Item" se estamos no resumo da SIN
        atuar = await page.evaluate("""() => {
            const btn = document.querySelector("input[value='Atuar no Item']");
            if (btn && !btn.disabled) { btn.click(); return true; }
            return false;
        }""")
        if atuar:
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
            await hide_overlays(page)

        # Navegar Descrições → Editar Descrição
        if "ITEM_Edita.aspx" in page.url:
            await page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
                const div1 = document.getElementById('div1');
                if (div1) div1.style.display = 'none';
                __doPostBack('ctl00$Body$dlTab$ctl01$lbutMenu', '');
            }""")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
            await hide_overlays(page)

            await page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
                __doPostBack('ctl00$Body$tabDescricoes$lbutAlterarDescr', '');
            }""")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)

        if "ITEM_Edita_DescricaoV3" in page.url and "IdItem=0" not in page.url:
            log.info(f"Re-navegação bem-sucedida: {page.url}")
        else:
            log.warning(f"Re-navegação falhou, URL atual: {page.url}")
            return False

    log.info(f"PDM alterado para: {pdm} / {found_cat}")
    return True
