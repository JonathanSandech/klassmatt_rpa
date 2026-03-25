"""Aba Referências — empresa e part number."""

import re

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, page_contains_text, wait_for_text
from logger import log


async def _get_ref_count(page: Page) -> int:
    """Retorna o número de referências pelo label da aba 'Referências (N)'."""
    tab_el = page.locator("a:has-text('Referências')").first
    try:
        tab_text = await tab_el.inner_text(timeout=5_000)
        match = re.search(r"\((\d+)\)", tab_text)
        if match:
            return int(match.group(1))
    except Exception:
        pass
    return 0


async def _read_existing_ref_text(page: Page) -> str:
    """Lê o texto 'Referência/Fabricante: XXX/YYY' da aba de referências.

    Usa a mesma lógica do verify — busca no innerText da página.
    Retorna o valor raw (ex: 'N/A/N/A' ou 'IS400TDBTH8A/BAKER HUGHES') ou ''.
    """
    return await page.evaluate("""() => {
        const allText = document.body.innerText || '';
        const matches = allText.match(/Referência\\/Fabricante:\\s*([^\\n]+)/g);
        if (matches && matches.length > 0) {
            // Retornar a primeira referência encontrada
            return matches[0].replace('Referência/Fabricante:', '').trim();
        }
        return '';
    }""")


def _is_placeholder_ref(raw: str) -> bool:
    """Retorna True se a referência é um placeholder (N/A, vazio, etc)."""
    if not raw:
        return True
    cleaned = raw.replace("/", "").replace("N", "").replace("A", "").replace(" ", "")
    return cleaned == "" or raw in ("N/A", "N/A/N/A", "/", "N")


async def _select_autocomplete(page: Page, empresa: str) -> bool:
    """Tenta selecionar a empresa no autocomplete. Retorna True se selecionou."""

    # 1. JS fallback — qualquer lista de autocomplete visível (mais confiável)
    try:
        clicked = await page.evaluate(
            """() => {
                const lists = document.querySelectorAll('.ac_results, .ui-autocomplete, [id*="autocomplete"]');
                for (const list of lists) {
                    if (list.offsetParent !== null) {
                        const items = list.querySelectorAll('li a, li');
                        if (items.length > 0) { items[0].click(); return true; }
                    }
                }
                return false;
            }"""
        )
        if clicked:
            return True
    except Exception:
        pass

    # 2. Match exato via texto (case-insensitive)
    try:
        await page.locator(f"a:text-matches('{re.escape(empresa)}', 'i')").first.click(timeout=3_000)
        return True
    except Exception:
        pass

    # 3. Qualquer item visível no autocomplete (jQuery UI, ac_results)
    for ac_sel in [
        ".ac_results li:first-child a",
        ".ac_results li:first-child",
        ".ui-autocomplete li:first-child a",
        ".ui-autocomplete li:first-child",
    ]:
        try:
            el = page.locator(ac_sel).first
            if await el.count() > 0 and await el.is_visible():
                await el.click(timeout=3_000)
                return True
        except Exception:
            continue

    # 4. Match parcial com primeira palavra
    first_word = empresa.split()[0] if empresa.split() else empresa
    try:
        await page.locator(f"a:text-matches('{re.escape(first_word)}', 'i')").first.click(timeout=5_000)
        return True
    except Exception:
        pass

    return False


async def _close_fabricante_tab_and_cancel_form(page: Page) -> None:
    """Fecha aba de cadastro de fabricante (se aberta) e cancela o form de edição.

    Quando o fabricante não existe, o confirm dialog "Deseja cadastra-lo?" abre
    FabricanteFornecManu.aspx numa nova aba. A referência já é salva nesse ponto,
    mas o form de edição continua aberto (dirty state). Precisamos:
    1. Fechar a aba do fabricante (se existir)
    2. Clicar Cancelar no form de edição para limpar o dirty state
    """
    # 1. Fechar abas auxiliares de fabricante (FabricanteFornecManu.aspx)
    try:
        all_pages = page.context.pages
        for p in all_pages:
            if p != page and "FabricanteFornecManu" in p.url:
                log.debug(f"Fechando aba de cadastro de fabricante: {p.url}")
                try:
                    await p.close()
                except Exception:
                    pass
    except Exception:
        pass

    await page.wait_for_timeout(500)

    # 2. Se o form de edição ainda está aberto, clicar Cancelar
    #    (a referência já foi salva — Cancelar apenas fecha o form)
    try:
        cancelar = page.locator("#btnCancelar").first
        if await cancelar.is_visible(timeout=1_500):
            log.debug("Form referência ainda aberto após save — clicando Cancelar para limpar dirty state")
            await page.evaluate("() => { const b = document.querySelector('#btnCancelar'); if (b) b.click(); }")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
    except Exception:
        pass

    # 3. Garantir que saímos do dirty state recarregando a aba
    try:
        await safe_click(page, SELECTORS["tab_referencias"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
    except Exception:
        pass


async def _verify_ref_saved(page: Page, part_number: str) -> bool:
    """Verifica se a referência foi realmente salva checando divReferencias."""
    try:
        ref_text = await page.evaluate("""() => {
            const div = document.querySelector('#divReferencias');
            return div ? div.innerText.trim() : '';
        }""")
        return part_number in ref_text
    except Exception:
        return False


async def fill_reference(page: Page, empresa: str, part_number: str) -> bool:
    """Preenche referência (empresa + part number).

    Retorna True se ok, False se referência duplicada.
    """
    log.info(f"Preenchendo referência: {empresa} / {part_number}")

    # Navegar para aba Referências
    await safe_click(page, SELECTORS["tab_referencias"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Verificar referência existente
    ref_count = await _get_ref_count(page)
    use_edit = False  # True = editar existente, False = adicionar nova

    if ref_count > 0:
        raw_ref = await _read_existing_ref_text(page)
        log.debug(f"  Referência existente raw: '{raw_ref}' (count={ref_count})")

        if _is_placeholder_ref(raw_ref):
            log.info(f"Referência existente é placeholder ('{raw_ref}') — editando")
            use_edit = True
        elif part_number in raw_ref:
            log.info(f"Referência já existe com part number correto ({part_number}) — pulando")
            return True
        else:
            log.info(f"Referência existente ('{raw_ref}') difere — editando existente")
            use_edit = True

    # Abrir formulário: EDIT (Imagebutton22) ou ADD (iButAddRef)
    # Usar JS evaluate para evitar form intercept de pointer events
    if use_edit:
        clicked = await page.evaluate("""() => {
            const btn = document.querySelector("input[type='image'][id$='Imagebutton22']");
            if (btn) { btn.click(); return true; }
            const add = document.querySelector('#iButAddRef');
            if (add) { add.click(); return true; }
            return false;
        }""")
        log.debug(f"Editando via {'Imagebutton22' if clicked else 'fallback'}")
    else:
        await page.evaluate("""() => {
            const add = document.querySelector('#iButAddRef');
            if (add) { add.click(); return; }
            const edit = document.querySelector("input[type='image'][id$='Imagebutton22']");
            if (edit) edit.click();
        }""")

    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Preencher empresa — usar press_sequentially para triggar autocomplete
    # (page.fill substitui tudo de vez e não gera os eventos de keystroke)
    # Triple-click+delete via JS para limpar campo corretamente (el.value='' não limpa binding)
    # Depois focus via JS para evitar intercept do aspnetForm overlay
    await page.evaluate("""() => {
        const el = document.querySelector('#txtNome');
        if (el) { el.select(); }
    }""")
    await page.keyboard.press("Delete")
    await page.wait_for_timeout(300)
    empresa_input = page.locator(SELECTORS["ref_empresa_input"])
    await empresa_input.press_sequentially(str(empresa), delay=50)
    await page.wait_for_timeout(2000)

    # Selecionar empresa no autocomplete
    autocomplete_ok = await _select_autocomplete(page, empresa)
    if not autocomplete_ok:
        log.warning(f"Autocomplete empresa '{empresa}' não encontrado — continuando sem seleção")

    # Preencher Part Number
    await safe_fill(page, SELECTORS["ref_partnumber_input"], str(part_number))

    # Salvar referência via btnSalvar (botão do form de referência, NÃO o butSalvar do footer)
    await page.evaluate("""() => {
        const btn = document.querySelector('#btnSalvar');
        if (btn) btn.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Verificar duplicidade (pode redirecionar para página de aviso)
    is_duplicate = await page_contains_text(page, SELECTORS["ref_duplicate_text"])
    if is_duplicate:
        log.warning(f"Referência duplicada detectada: {empresa} / {part_number}")
        try:
            await page.evaluate("""() => {
                const btn = document.querySelector("input[value='Continuar']");
                if (btn) btn.click();
            }""")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1000)
        except Exception:
            pass
        return False

    # Fechar aba de fabricante se abriu + cancelar form se ainda aberto
    await _close_fabricante_tab_and_cancel_form(page)

    # Recarregar aba referências para garantir estado limpo antes de verificar
    try:
        await safe_click(page, SELECTORS["tab_referencias"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
    except Exception:
        pass

    # Verificar se a referência realmente foi salva no Klassmatt
    saved = await _verify_ref_saved(page, part_number)
    if saved:
        log.info("Referência salva com sucesso")
    else:
        log.warning(f"Referência NÃO salvou (part_number '{part_number}' não encontrado em divReferencias)")
    return saved
