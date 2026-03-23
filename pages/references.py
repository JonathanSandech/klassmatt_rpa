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

    # 1. Match exato via texto
    try:
        await page.locator(f"a:has-text('{empresa}')").first.click(timeout=3_000)
        return True
    except Exception:
        pass

    # 2. Qualquer item visível no autocomplete (jQuery UI, ac_results)
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

    # 3. JS fallback — qualquer lista de autocomplete visível
    try:
        clicked = await page.evaluate(
            """() => {
                const lists = document.querySelectorAll('.ac_results, .ui-autocomplete, [id*="autocomplete"]');
                for (const list of lists) {
                    if (list.offsetParent !== null) {
                        const items = list.querySelectorAll('li, a');
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

    # 4. Match parcial com primeira palavra
    first_word = empresa.split()[0] if empresa.split() else empresa
    try:
        await page.locator(f"a:text-matches('{first_word}', 'i')").first.click(timeout=5_000)
        return True
    except Exception:
        pass

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
            log.info(f"Referência existente ('{raw_ref}') difere — adicionando nova")
            use_edit = False

    # Abrir formulário: EDIT (Imagebutton22) ou ADD (iButAddRef)
    if use_edit:
        edit_btn = page.locator("input[type='image'][id$='Imagebutton22']").first
        if await edit_btn.count() > 0:
            log.debug("Editando via Imagebutton22")
            await edit_btn.click()
        else:
            log.warning("Imagebutton22 não encontrado — usando ADD")
            await page.locator("#iButAddRef").click()
    else:
        add_btn = page.locator("#iButAddRef")
        if await add_btn.count() > 0 and await add_btn.is_visible():
            await add_btn.click()
        else:
            edit_btn = page.locator("input[type='image'][id$='Imagebutton22']")
            await edit_btn.click()

    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Preencher empresa (digitar para acionar autocomplete)
    await safe_fill(page, SELECTORS["ref_empresa_input"], str(empresa))
    await page.wait_for_timeout(2000)

    # Selecionar empresa no autocomplete
    autocomplete_ok = await _select_autocomplete(page, empresa)
    if not autocomplete_ok:
        log.warning(f"Autocomplete empresa '{empresa}' não encontrado — continuando sem seleção")

    # Preencher Part Number
    await safe_fill(page, SELECTORS["ref_partnumber_input"], str(part_number))

    # Salvar — o confirm dialog "Fabricante não existe, deseja cadastrá-lo?" é aceito
    # automaticamente pelo handler global.
    salvar_btn = page.locator(SELECTORS["ref_salvar_btn"]).first
    await salvar_btn.click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    # Verificar duplicidade (pode redirecionar para página de aviso)
    is_duplicate = await page_contains_text(page, SELECTORS["ref_duplicate_text"])
    if is_duplicate:
        log.warning(f"Referência duplicada detectada: {empresa} / {part_number}")
        # Clicar Continuar se disponível
        continuar = page.locator("input[value='Continuar']")
        try:
            if await continuar.count() > 0:
                await continuar.click()
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)
        except Exception:
            pass
        return False

    # Após salvar com fabricante novo, o form pode continuar aberto (dirty state).
    # Tentar salvar novamente se o botão salvar ainda estiver visível.
    for retry in range(2):
        try:
            salvar_still = page.locator(SELECTORS["ref_salvar_btn"]).first
            if await salvar_still.is_visible(timeout=1_500):
                log.debug(f"Form referência ainda aberto (retry {retry + 1}) — salvando novamente")
                await salvar_still.click()
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(1000)
            else:
                break
        except Exception:
            break

    # Garantir que o form não está em dirty state antes de sair da aba.
    # Recarregar a aba de referências para limpar qualquer estado pendente.
    try:
        await safe_click(page, SELECTORS["tab_referencias"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)
    except Exception:
        pass

    log.info("Referência salva com sucesso")
    return True
