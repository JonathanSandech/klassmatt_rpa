"""Aba Relacionamentos — CÓDIGO ANTIGO, ATIVO ERP, ZBRA."""

from playwright.async_api import Page

import browser as _browser
from config import SELECTORS, RELATIONSHIP_TYPE, RELATIONSHIP_STATUS, RELATIONSHIP_COMMENT
from browser import safe_click, safe_fill
from logger import log


async def fill_relationship(page: Page, codigo_60: str) -> None:
    """Preenche relacionamento com código antigo.

    Tipo: CÓDIGO ANTIGO
    Código: valor de 'Código 60' do Excel
    Status: ATIVO ERP
    Comentário: ZBRA
    """
    log.info(f"Preenchendo relacionamento: {codigo_60}")

    # Navegar para aba Relacionamentos
    # O ASP.NET pode bloquear se houver alterações pendentes em outra aba;
    # o dialog handler aceita o alert, mas precisamos tentar novamente.
    for tab_attempt in range(3):
        await safe_click(page, SELECTORS["tab_relacionamentos"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)

        # Verificar se realmente estamos na aba de Relacionamentos
        add_btn = page.locator(SELECTORS["rel_add_btn"])
        if await add_btn.count() > 0:
            break
        log.debug(f"Aba Relacionamentos não carregou (tentativa {tab_attempt + 1}/3)")
    else:
        raise RuntimeError("Não conseguiu navegar para aba Relacionamentos após 3 tentativas")

    # Clicar no botão de adicionar
    await add_btn.click()
    await page.wait_for_load_state("networkidle")

    # Selecionar tipo: CÓDIGO ANTIGO
    await safe_click(page, SELECTORS["rel_tipo_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_TYPE}')")

    # Preencher código
    await safe_fill(page, SELECTORS["rel_codigo_input"], str(codigo_60))

    # Selecionar status: ATIVO ERP
    await safe_click(page, SELECTORS["rel_status_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_STATUS}')")

    # Preencher comentário: ZBRA
    await safe_fill(page, SELECTORS["rel_comentario_input"], RELATIONSHIP_COMMENT)

    # Limpar flag de dialog antes de salvar
    _browser.last_dialog_message = ""

    # Salvar relacionamento
    save_btn = page.locator(SELECTORS["rel_save_btn"])
    await save_btn.click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Verificar se o relacionamento já existia (alert "Este código já está relacionado...")
    last_msg = _browser.last_dialog_message.lower()
    if "já está relacionado" in last_msg or "already" in last_msg:
        log.info(f"Relacionamento já existia: {RELATIONSHIP_TYPE} / {codigo_60} — pulando")
    else:
        log.info(f"Relacionamento salvo: {RELATIONSHIP_TYPE} / {codigo_60}")
