"""Aba Classificações — UNSPSC.

Performance notes (2026-03):
    - UNSPSC was the #1 bottleneck (~45-60s per item, 45-50% of total time).
    - ASP.NET postbacks require networkidle waits (unavoidable).
    - Explicit wait_for_timeout calls were reduced from 3.5s total to 0.5s.
    - Each sub-step is timed so regressions can be identified in the logs.
"""

import time

from playwright.async_api import Page

from config import SELECTORS
from browser import safe_click, safe_fill, hide_overlays
from logger import log


async def fill_unspsc(page: Page, unspsc_code: str) -> bool:
    """Preenche o código UNSPSC na aba Classificações.

    Sequência: Aba Classificações → botão UNSPSC → preencher código →
    Pesquisar → verificar resultado → selecionar → Selecionar

    Se o código não existir no Klassmatt, loga erro e não seleciona nada.
    """
    t0 = time.perf_counter()
    log.info(f"Preenchendo UNSPSC: {unspsc_code}")

    # --- 1. Navegar para aba Classificações (postback) ---
    t1 = time.perf_counter()
    await safe_click(page, SELECTORS["tab_classificacoes"])
    await page.wait_for_load_state("networkidle")  # postback — must wait
    await hide_overlays(page)
    log.debug(f"  UNSPSC [tab_classificacoes] {time.perf_counter() - t1:.1f}s")

    # --- 2. Idempotency check — fast JS evaluate, always run ---
    t1 = time.perf_counter()
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
    log.debug(f"  UNSPSC [idempotency_check] {time.perf_counter() - t1:.1f}s")
    if current_unspsc and current_unspsc.replace('.', '') == str(unspsc_code).strip():
        log.info(f"UNSPSC já preenchido corretamente ({current_unspsc}) — pulando")
        return True

    # --- 3. Abrir popup UNSPSC (postback) ---
    t1 = time.perf_counter()
    await safe_click(page, SELECTORS["unspsc_btn"])
    await page.wait_for_load_state("networkidle")  # postback — must wait
    # Small buffer for ASP.NET panel to render after postback completes
    await page.wait_for_timeout(200)
    log.debug(f"  UNSPSC [open_popup] {time.perf_counter() - t1:.1f}s")

    # --- 4. Preencher código ---
    await safe_fill(page, SELECTORS["unspsc_input"], str(unspsc_code))

    # --- 5. Pesquisar (postback) ---
    t1 = time.perf_counter()
    await safe_click(page, SELECTORS["unspsc_pesquisar_btn"])
    await page.wait_for_load_state("networkidle")  # postback — must wait
    # No extra wait needed: networkidle already guarantees DOM is stable
    log.debug(f"  UNSPSC [pesquisar] {time.perf_counter() - t1:.1f}s")

    # --- 6. Verificar resultado ---
    t1 = time.perf_counter()
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
    log.debug(f"  UNSPSC [verify_result] {time.perf_counter() - t1:.1f}s")

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
        return False

    # --- 7. Marcar checkbox (postback — input[type="image"] needs real click) ---
    t1 = time.perf_counter()
    await page.click("input[id$='ckSelUNSPSC']")
    await page.wait_for_load_state("networkidle")  # postback — must wait
    # No extra wait: networkidle is sufficient for the postback to complete
    log.debug(f"  UNSPSC [checkbox] {time.perf_counter() - t1:.1f}s")

    # --- 8. Confirmar seleção (postback) ---
    t1 = time.perf_counter()
    await safe_click(page, SELECTORS["unspsc_selecionar_btn"])
    await page.wait_for_load_state("networkidle")  # postback — must wait
    log.debug(f"  UNSPSC [selecionar] {time.perf_counter() - t1:.1f}s")

    # --- 9. Salvar (postback — sem Salvar, valor é perdido ao trocar de aba) ---
    t1 = time.perf_counter()
    await safe_click(page, SELECTORS["salvar_btn"])
    await page.wait_for_load_state("networkidle")  # postback — must wait
    # Small buffer to ensure save confirmation is rendered
    await page.wait_for_timeout(300)
    log.debug(f"  UNSPSC [salvar] {time.perf_counter() - t1:.1f}s")

    elapsed = time.perf_counter() - t0
    log.info(f"UNSPSC {unspsc_code} selecionado e salvo ({elapsed:.1f}s)")
    return True
