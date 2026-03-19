"""Aba Relacionamentos — CÓDIGO ANTIGO, ATIVO ERP, ZBRA."""

from playwright.async_api import Page

import browser as _browser
from config import SELECTORS, RELATIONSHIP_TYPE, RELATIONSHIP_STATUS, RELATIONSHIP_COMMENT
from browser import safe_click, safe_fill
from logger import log


async def _get_existing_relationships(page: Page) -> list[dict]:
    """Lê relacionamentos existentes da grid #dgRelacionamento via JS evaluate.

    Retorna lista de dicts com keys: tipo, codigo, status, comentario, row_index.
    Layout da grid (confirmado via MCP):
      Row 0 = header: Relacionamento | Código | Status | Observações | [Imagebutton7=ADD]
      Rows 1..N = dados: tipo | código | status | obs | [ibutEditRelac=EDIT]
      Última row = vazia (formulário de edição inline)
    """
    rels = await page.evaluate("""() => {
        const table = document.querySelector('#dgRelacionamento');
        if (!table) return [];
        const result = [];
        for (let i = 1; i < table.rows.length; i++) {
            const cells = table.rows[i].querySelectorAll('td');
            if (cells.length < 4) continue;
            const tipo = (cells[0]?.innerText || '').trim();
            const codigo = (cells[1]?.innerText || '').trim();
            const status = (cells[2]?.innerText || '').trim();
            const comentario = (cells[3]?.innerText || '').trim();
            // Pular rows vazias (última row é o formulário inline)
            if (!tipo && !codigo) continue;
            const editBtn = cells[4]?.querySelector('input[id="ibutEditRelac"]');
            result.push({
                tipo, codigo, status, comentario,
                row_index: i,
                edit_btn_name: editBtn ? editBtn.name : null
            });
        }
        return result;
    }""")
    return rels or []


async def _navigate_to_tab(page: Page) -> None:
    """Navega para aba Relacionamentos com retry."""
    for tab_attempt in range(3):
        await safe_click(page, SELECTORS["tab_relacionamentos"])
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)

        add_btn = page.locator(SELECTORS["rel_add_btn"])
        if await add_btn.count() > 0:
            return
        log.debug(f"Aba Relacionamentos não carregou (tentativa {tab_attempt + 1}/3)")
    raise RuntimeError("Não conseguiu navegar para aba Relacionamentos após 3 tentativas")


async def _add_relationship(page: Page, codigo_60: str) -> None:
    """Adiciona um novo relacionamento (Adicionar → preencher → Salvar)."""
    add_btn = page.locator(SELECTORS["rel_add_btn"])
    await add_btn.click()
    await page.wait_for_load_state("networkidle")

    await _fill_fields(page, codigo_60)
    await _save(page, codigo_60)


async def _fill_fields(page: Page, codigo_60: str) -> None:
    """Preenche os campos do formulário de relacionamento."""
    # Tipo: CÓDIGO ANTIGO
    await safe_click(page, SELECTORS["rel_tipo_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_TYPE}')")

    # Código
    await safe_fill(page, SELECTORS["rel_codigo_input"], str(codigo_60))

    # Status: ATIVO ERP
    await safe_click(page, SELECTORS["rel_status_input"])
    await page.wait_for_timeout(500)
    await safe_click(page, f"a:has-text('{RELATIONSHIP_STATUS}')")

    # Comentário: ZBRA
    await safe_fill(page, SELECTORS["rel_comentario_input"], RELATIONSHIP_COMMENT)


async def _save(page: Page, codigo_60: str) -> None:
    """Salva o relacionamento e verifica resultado."""
    _browser.last_dialog_message = ""

    save_btn = page.locator(SELECTORS["rel_save_btn"])
    await save_btn.click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    last_msg = _browser.last_dialog_message.lower()
    if "já está relacionado" in last_msg or "already" in last_msg:
        log.info(f"Relacionamento já existia: {RELATIONSHIP_TYPE} / {codigo_60} — pulando")
    else:
        log.info(f"Relacionamento salvo: {RELATIONSHIP_TYPE} / {codigo_60}")


async def fill_relationship(page: Page, codigo_60: str) -> None:
    """Preenche relacionamento com código antigo.

    Tipo: CÓDIGO ANTIGO
    Código: valor de 'Código 60' do Excel
    Status: ATIVO ERP
    Comentário: ZBRA

    Verifica relacionamentos existentes antes de adicionar:
    - Se já existe CÓDIGO ANTIGO com o mesmo código → skip
    - Se já existe CÓDIGO ANTIGO com código diferente → edita para o novo
    - Se não existe CÓDIGO ANTIGO → adiciona novo
    """
    log.info(f"Preenchendo relacionamento: {codigo_60}")

    await _navigate_to_tab(page)

    # Ler relacionamentos existentes da grid
    existing = await _get_existing_relationships(page)
    if existing:
        log.debug(f"Relacionamentos existentes: {len(existing)} encontrados")
        for rel in existing:
            log.debug(f"  → {rel['tipo']} / {rel['codigo']} / {rel['status']} / {rel['comentario']}")

    # Procurar relacionamento do tipo CÓDIGO ANTIGO
    matching = [r for r in existing if r["tipo"].upper() == RELATIONSHIP_TYPE.upper()]

    if matching:
        existing_rel = matching[0]
        existing_code = existing_rel["codigo"].strip()
        target_code = str(codigo_60).strip()

        if existing_code == target_code:
            log.info(f"Relacionamento já existe com mesmo código ({target_code}) — pulando")
            return

        # Código diferente → editar o existente
        log.info(f"Relacionamento existente com código diferente: {existing_code} → {target_code}")
        log.info("Editando relacionamento existente...")

        # Clicar no botão de editar via name (único por row)
        edit_name = existing_rel.get("edit_btn_name")
        if edit_name:
            edited = await page.evaluate("""(btnName) => {
                const btn = document.querySelector('input[name="' + btnName + '"]');
                if (btn) { btn.click(); return true; }
                return false;
            }""", edit_name)
        else:
            edited = False

        if edited:
            await page.wait_for_load_state("networkidle")
            await _fill_fields(page, codigo_60)
            await _save(page, codigo_60)
        else:
            log.warning("Não encontrou botão de editar — adicionando novo relacionamento")
            await _add_relationship(page, codigo_60)
    else:
        # Nenhum CÓDIGO ANTIGO existe → adicionar novo
        await _add_relationship(page, codigo_60)
