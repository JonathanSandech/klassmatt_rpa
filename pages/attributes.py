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


async def fill_attributes(page: Page, attributes: list) -> bool:
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
            return False

    # Verificar se a tabela de atributos existe
    has_grid = await page.locator("#dgDadosTecnicos").count() > 0
    if not has_grid:
        log.warning("Tabela dgDadosTecnicos não encontrada — pulando atributos")
        return False

    # Contar quantos atributos existem na grid
    total_attrs = await page.evaluate(
        """() => {
            const dg = document.querySelector('#dgDadosTecnicos');
            if (!dg) return 0;
            return dg.querySelectorAll('tr').length - 1;  // -1 para header
        }"""
    )

    for i in range(total_attrs):
        ctl_idx = _attr_ctl_index(i + 1)

        # Pegar valor da planilha (ou N/A se acabaram os valores)
        value = attributes[i] if i < len(attributes) else None
        if value is None or (isinstance(value, str) and value.strip() == ""):
            value_str = "N/A"
        else:
            value_str = str(value).strip()

        # Verificar se o atributo já está preenchido
        current_val = await page.evaluate(
            f"""() => {{
                const row = document.querySelector("input[name$='dgDadosTecnicos$ctl{ctl_idx}$hdnDtTexto']");
                return row ? row.value.trim() : '';
            }}"""
        )
        if current_val and current_val != "":
            # Comparar com o valor da planilha — pular só se bate
            if current_val.upper() == value_str.upper():
                log.debug(f"Atributo {i + 1}: já preenchido com '{current_val}' — pulando")
                continue
            elif current_val.upper() == "N/A" and value_str.upper() != "N/A":
                log.info(f"Atributo {i + 1}: existente='N/A' mas planilha='{value_str}' — sobrescrevendo")
                # Precisa limpar N/A antes de preencher — desmarcar checkbox N/A
                na_selector = SELECTORS["attr_na_checkbox_tpl"].format(idx=ctl_idx)
                na_el = page.locator(na_selector)
                if await na_el.count() > 0:
                    try:
                        if await na_el.is_checked():
                            await safe_click(page, na_selector)
                            await page.wait_for_timeout(1000)
                    except Exception:
                        pass
            else:
                log.debug(f"Atributo {i + 1}: já preenchido com '{current_val}' (planilha='{value_str}') — pulando")
                continue

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

            # Abrir popup da árvore via AbreJanTaxonomia() e preencher (com retry)
            for popup_attempt in range(2):
                try:
                    await _open_and_fill_tree_popup(page, ctl_idx, value_str)
                    break
                except Exception as popup_err:
                    if popup_attempt == 0:
                        log.warning(f"Atributo {i + 1}: popup falhou ({popup_err}) — retentando")
                        await page.wait_for_timeout(2000)
                    else:
                        log.warning(f"Atributo {i + 1}: popup falhou 2x — pulando atributo")

            # Após popup fechar, verificar se ainda estamos na DescricaoV3
            await page.wait_for_timeout(1000)
            if "ITEM_Edita_DescricaoV3" not in page.url:
                log.debug(f"Página mudou após popup: {page.url} — re-navegando para DescricaoV3")
                if "ITEM_Resumo" in page.url or "SIN_Item_Resultante" in page.url:
                    atuar = page.locator(SELECTORS["atuar_no_item_btn"])
                    if await atuar.count() > 0:
                        await atuar.click()
                        await page.wait_for_load_state("networkidle")
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

    # Finalizar para persistir os atributos (sem Finalizar, valores são perdidos)
    _finalizar_ok = True
    if "ITEM_Edita_DescricaoV3" in page.url:
        import browser as _browser
        _browser.last_dialog_message = ""

        finalizar_btn = page.locator("#butFinaliza")
        if await finalizar_btn.count() > 0:
            log.debug("Clicando Finalizar na DescricaoV3 para salvar atributos...")
            try:
                await finalizar_btn.click()
                await page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                try:
                    await page.wait_for_load_state("networkidle", timeout=15_000)
                except Exception:
                    pass
            await page.wait_for_timeout(1000)

            # Verificar se Finalizar foi rejeitado (alert de dados técnicos incompletos)
            if "ITEM_Edita_DescricaoV3" in page.url:
                last_msg = _browser.last_dialog_message.lower()
                if "preencher" in last_msg or "verificar" in last_msg or "dados técnicos" in last_msg:
                    log.warning(f"Finalizar rejeitado: '{_browser.last_dialog_message}' — atributos não salvos")
                    _finalizar_ok = False
                else:
                    log.warning("Ainda na DescricaoV3 após Finalizar — pode não ter salvo")
                    _finalizar_ok = False
            else:
                log.debug("Atributos finalizados e salvos com sucesso")

    # Voltar para a página do item (ITEM_Edita.aspx) se ainda estamos na DescricaoV3
    if "ITEM_Edita_DescricaoV3" in page.url:
        log.debug("Voltando da DescricaoV3 para a página do item...")
        voltar_btn = page.locator("#butSIN_Voltar")
        if await voltar_btn.count() > 0:
            await voltar_btn.click()
            await page.wait_for_load_state("networkidle", timeout=15_000)
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
    return _finalizar_ok


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

        # Esperar que a árvore renderize seus nós (ASP.NET TreeView usa JS client-side)
        try:
            await popup_page.wait_for_selector(
                "a.nodeStyle, a.nodeStyleSel, a[class*='nodeStyle']",
                timeout=10_000,
            )
        except Exception:
            log.warning("Nós da árvore não apareceram após 10s — tentando mesmo assim")
        await popup_page.wait_for_timeout(500)

        first_letter = value[0].upper()

        # Passo 1: Clicar na letra do alfabeto para expandir a sub-árvore
        # As letras são links com __doPostBack que causam navegação/postback
        # O click causa full page reload — precisamos esperar a nova página carregar
        letter_found = await popup_page.evaluate(
            """(letter) => {
                // Tentar múltiplos seletores para encontrar a letra
                const selectors = [
                    'a.nodeStyle', 'a.nodeStyleSel',
                    'a[class*="nodeStyle"]', 'a[class*="NodeStyle"]'
                ];
                for (const sel of selectors) {
                    const nodes = document.querySelectorAll(sel);
                    const letterNode = Array.from(nodes).find(a => a.innerText.trim() === letter);
                    if (letterNode) {
                        letterNode.click();
                        return true;
                    }
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
            # Esperar nós filhos renderizarem após o postback
            await popup_page.wait_for_timeout(2000)
        else:
            # Listar o que está disponível para debug
            available = await popup_page.evaluate(
                """() => {
                    const nodes = document.querySelectorAll('a.nodeStyle, a.nodeStyleSel, a[class*="nodeStyle"]');
                    return Array.from(nodes).map(a => a.innerText.trim()).slice(0, 30);
                }"""
            )
            log.warning(f"Letra '{first_letter}' não encontrada na árvore. Nós disponíveis: {available}")

        # Passo 2: Procurar o valor nos nós expandidos (com fuzzy matching)
        # Usar evaluate pois pode haver milhares de nós (1900+)
        found = await popup_page.evaluate(
            """(value) => {
                const nodes = document.querySelectorAll('a.nodeStyle, a.nodeStyleSel');
                const upper = value.toUpperCase().trim();

                // 1. Exact match
                let target = Array.from(nodes).find(a => a.innerText.trim() === value);
                let matchType = 'exact';

                // 2. Case-insensitive exact
                if (!target) {
                    target = Array.from(nodes).find(a => a.innerText.trim().toUpperCase() === upper);
                    matchType = 'case-insensitive';
                }

                // 3. "starts with" match — tree value starts with Excel value or vice versa
                if (!target) {
                    target = Array.from(nodes).find(a => {
                        const nodeText = a.innerText.trim().toUpperCase();
                        return nodeText.startsWith(upper) || upper.startsWith(nodeText);
                    });
                    matchType = 'starts-with';
                }

                // 4. "contains all words" match — all words from Excel value appear in tree node
                if (!target) {
                    const words = upper.split(/\\s+/).filter(w => w.length > 2);
                    if (words.length > 0) {
                        target = Array.from(nodes).find(a => {
                            const nodeText = a.innerText.trim().toUpperCase();
                            return words.every(w => nodeText.includes(w));
                        });
                        matchType = 'contains-all-words';
                    }
                }

                // 5. "most words match" — find node with most matching words (minimum 60%)
                if (!target) {
                    const words = upper.split(/\\s+/).filter(w => w.length > 2);
                    if (words.length > 0) {
                        let bestMatch = null;
                        let bestScore = 0;
                        for (const node of nodes) {
                            const nodeText = node.innerText.trim().toUpperCase();
                            const matchCount = words.filter(w => nodeText.includes(w)).length;
                            const score = matchCount / words.length;
                            if (score > bestScore && score >= 0.6) {
                                bestScore = score;
                                bestMatch = node;
                            }
                        }
                        if (bestMatch) {
                            target = bestMatch;
                            matchType = 'best-word-match(' + Math.round(bestScore * 100) + '%)';
                        }
                    }
                }

                if (target) {
                    target.click();
                    return { found: true, text: target.innerText.trim(), matchType: matchType };
                }
                return { found: false };
            }""",
            value,
        )

        if found and found.get("found"):
            matched_text = found.get("text", value)
            match_type = found.get("matchType", "unknown")
            if matched_text != value:
                log.info(f"Atributo fuzzy match: '{value}' -> '{matched_text}' ({match_type})")
            else:
                log.debug(f"Nó selecionado na árvore: '{matched_text}' ({match_type})")
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
