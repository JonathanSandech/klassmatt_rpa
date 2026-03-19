"""Aba Classificações — UNSPSC."""

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, hide_overlays
from logger import log


async def fill_unspsc(page: Page, unspsc_code: str) -> None:
    """Preenche o código UNSPSC na aba Classificações.

    Sequência: Aba Classificações → botão UNSPSC → preencher código →
    Pesquisar → verificar resultado → selecionar → Selecionar

    Se o código não existir no Klassmatt, loga erro e não seleciona nada.
    """
    log.info(f"Preenchendo UNSPSC: {unspsc_code}")

    # Navegar para aba Classificações
    await safe_click(page, SELECTORS["tab_classificacoes"])
    await page.wait_for_load_state("networkidle")
    await hide_overlays(page)

    # Verificar se UNSPSC já está preenchido com o valor correto (idempotente)
    current_unspsc = await page.evaluate(
        """() => {
            const el = document.querySelector('#txtUNSPSC') ||
                       document.querySelector('input[id*="UNSPSC"]');
            if (el && el.value) return el.value.trim();
            // Tentar pegar do texto visível na aba
            const spans = document.querySelectorAll('span, td');
            for (const s of spans) {
                const m = s.innerText.match(/^(\\d{8})\\./);
                if (m) return m[0].replace('.', '');
            }
            return '';
        }"""
    )
    if current_unspsc and current_unspsc.replace('.', '') == str(unspsc_code).strip():
        log.info(f"UNSPSC já preenchido corretamente ({current_unspsc}) — pulando")
        return

    # Clicar no botão UNSPSC
    await safe_click(page, SELECTORS["unspsc_btn"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Preencher código
    await safe_fill(page, SELECTORS["unspsc_input"], str(unspsc_code))

    # Pesquisar
    await safe_click(page, SELECTORS["unspsc_pesquisar_btn"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Verificar se a pesquisa retornou resultados REAIS (não "Nenhum registro")
    search_result = await page.evaluate(
        """(expectedCode) => {
            const noResult = document.body.innerText.includes('Nenhum registro');
            if (noResult) return { found: false, reason: 'nenhum_registro' };

            const checkboxes = document.querySelectorAll("input[id$='ckSelUNSPSC']");
            if (checkboxes.length === 0) return { found: false, reason: 'sem_checkbox' };

            // Procurar o código exato na grid de resultados
            const rows = document.querySelectorAll('tr');
            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                for (const cell of cells) {
                    if (cell.innerText.trim() === expectedCode) {
                        return { found: true, code: expectedCode };
                    }
                }
            }

            // Se não encontrou exato, verificar primeiro resultado
            const firstRow = checkboxes[0].closest('tr');
            const firstCode = firstRow ? firstRow.querySelector('td') : null;
            const firstCodeText = firstCode ? firstCode.innerText.trim() : '';
            return { found: false, reason: 'codigo_diferente', firstResult: firstCodeText };
        }""",
        str(unspsc_code),
    )

    if not search_result.get("found"):
        reason = search_result.get("reason", "desconhecido")
        first = search_result.get("firstResult", "")
        log.warning(
            f"UNSPSC {unspsc_code} não encontrado no Klassmatt "
            f"(motivo: {reason}, primeiro resultado: '{first}'). "
            f"Cancelando seleção."
        )
        cancel_btn = page.locator("input[value='Cancelar']")
        try:
            if await cancel_btn.count() > 0:
                await cancel_btn.click()
                await page.wait_for_load_state("networkidle")
        except Exception:
            pass
        return

    # Marcar checkbox via Playwright click (input[type="image"] precisa de click real
    # com coordenadas x,y para disparar o postback ASP.NET — JS cb.click() não funciona)
    await page.click("input[id$='ckSelUNSPSC']")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(500)

    # Confirmar seleção
    await safe_click(page, SELECTORS["unspsc_selecionar_btn"])
    await page.wait_for_load_state("networkidle")

    # Salvar para persistir o UNSPSC (sem Salvar, valor é perdido ao trocar de aba)
    await safe_click(page, SELECTORS["salvar_btn"])
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    log.info(f"UNSPSC {unspsc_code} selecionado e salvo")
