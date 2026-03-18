# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

RPA bot for automated bulk item registration in the Klassmatt MODEC web system (ASP.NET). Recently migrated from Power Automate Desktop to Python. Uses Playwright for browser automation with a persistent browser profile for session management.

## Infrastructure — Multi-VM Pipeline

The Klassmatt registration process runs across **4 VMs in parallel**, each handling a different pipeline stage. This bot (VM `.110`) executes the **final stage** — the actual item registration in Klassmatt.

### Shared Directory

All VMs read/write to a shared network folder:

```
\\WKS-TESTUSER5\Users\sandechin\Desktop\MODEC SHARED\
├── downloads/                          # PDFs downloaded by upstream VMs (~693 files)
├── documentos_baixados_<VM>.xlsx       # Document download manifest per VM
│   Columns: PART NUMBER, TAG, Nome do Documento, Data do Documento, Caminho
├── ptC_resultado_<VM>.xlsx             # Processing results per VM
│   Columns: PART NUMBER, G09, Tempo (s)
```

- `<VM>` = VM identifier (111, 112, etc.)
- The auxiliary spreadsheets are **informational only** — they track what upstream VMs have downloaded/processed
- `DOCUMENTS_DIR` should point to `\\WKS-TESTUSER5\Users\sandechin\Desktop\MODEC SHARED\downloads` for production runs
- Currently the bot only **reads** from the shared directory; write support will be added later

## Commands

```bash
# Install dependencies
pip install -r requirements.txt
python -m playwright install chromium

# Run the bot
python main.py
```

First run requires manual login in the browser window; the session is saved to `./playwright_profile/`. On subsequent runs, the bot resumes from `progress.json`, skipping items with status "ok".

There are no automated tests.

## Architecture

**Async orchestrator pattern** — `main.py` drives a 12-step workflow per Excel row through Page Object modules.

### Core modules (root level)

- **main.py** — Entry point and orchestrator. `run()` loads Excel, validates documents, launches browser, iterates items with retry logic (3 attempts + backoff). `process_item()` calls page objects in sequence. Includes `StepTimer` class for per-step timing instrumentation.
- **config.py** — All configuration: URLs, timeouts, Excel column mappings, and CSS selectors for ASP.NET elements. Also loads `.env`.
- **browser.py** — Playwright persistent context setup + helper functions (`safe_click`, `safe_fill`, `wait_for_text`, `retry_action`, session verification, popup dismissal, dialog auto-accept).
- **excel_handler.py** — Reads Excel via openpyxl, auto-detects header row, maps columns including 30 dynamic attributes (Atrib_1_Valor..Atrib_30_Valor). Colors rows green/red/orange for success/error/skip. `validate_documents()` resolves doc paths against `DOCUMENTS_DIR`.
- **state.py** — JSON-based progress persistence (`progress.json`). Tracks item status: "ok", "error", "skipped", "duplicate".
- **logger.py** — Dual output: file at DEBUG (`klassmatt_rpa.log`), console at INFO.

### Page Objects (`pages/`)

Each module encapsulates interaction with a specific Klassmatt UI section:

| Module | Step |
|---|---|
| `item.py` | Search SIN via `#DIVResultado` grid, select via `abreSIN()` JS, create item (skips if already created), finalize & remit to MODEC (adapts flow to available buttons) |
| `classifications.py` | UNSPSC code via popup search/select (`#ckSelUNSPSC`) |
| `fiscal.py` | NCM field — validates after fill, clears if rejected by ASP.NET |
| `references.py` | Company reference + part number; detects duplicate errors |
| `relationships.py` | Old code relationship (Type=CÓDIGO ANTIGO, Status=ATIVO ERP, Comment=ZBRA) |
| `media.py` | Document upload — detects new browser tab (`Midia.aspx`) for media page |
| `descriptions.py` | SAP description validation (reads `#txtD2` for length check, "Exibe D2" toggle), PDM category change |
| `attributes.py` | Up to 30 technical attributes — opens `Dt_EditaArvore.aspx` popup per attribute, navigates alphabet tree, selects value |
| `worklist.py` | Navigate to Worklist, filter "Todas as Solicitações" via select2 + JS `pesquisar()` |

## Key Patterns

- **All automation is async/await** — Playwright async API throughout.
- **Selectors are centralized in `config.py`** — never hardcode selectors in page objects; add them to `config.SELECTORS`.
- **ASP.NET HTML IDs differ from `name` attributes** — The `name` uses hierarchy like `ctl00$Body$tabFiscal$txtNCMTIPI` but the rendered `id` is just `txtNCMTIPI`. Always use the actual `id` attribute for selectors, not the `name`. IDs are also case-sensitive (e.g., `Imagebutton7` not `imagebutton7`).
- **Worklist grid uses `<div>` not `<table>`** — Results are in `#DIVResultado > .result` divs, not `table.GridClass` rows. SIN links use `javascript:abreSIN({id})`.
- **Select2 dropdowns** — The worklist filter is a select2 widget wrapping a native `<select>`. `page.select_option()` changes the value but doesn't trigger the JS search; call `pesquisar(0, '')` via `page.evaluate()` after.
- **Mídias tab opens a new browser tab** — The link uses `OpenNewTab()` JS. Detect the new tab by checking `page.context.pages` for `Midia.aspx` in URL.
- **NCM validation fires alerts on blur** — When an invalid NCM is entered, ASP.NET fires cascading `alert()` dialogs when focus leaves the field or when switching tabs. The field must be cleared if rejected to prevent blocking subsequent interactions. The bot also detects readonly NCM fields (items partially processed) and skips them.
- **Items may already be created** — SINs in FINALIZACAO status already have an item. The bot detects this (no "Criar item" button) and skips the creation step.
- **Attribute tree popup (`Dt_EditaArvore.aspx`)** — The `dgDadosTecnicos` table lives in `ITEM_Edita_DescricaoV3.aspx` (not `ITEM_Edita.aspx`). Each attribute's edit button (`btnAddEdit`) calls `AbreJanTaxonomia()` which opens a popup via `window.open()`. The tree has a hierarchical structure: root node → alphabet letters (A-Z) → values. Clicking a letter triggers `__doPostBack` (full page reload in popup). Use JS `evaluate` to click the letter, then `wait_for_load_state("networkidle")` to wait for the postback. Use `evaluate` to find/click the value among potentially 1900+ nodes, then Playwright `click()` for the "Selecionar" button.
- **Attribute selectors use `name`, not `id`** — In `dgDadosTecnicos`, all `btnAddEdit` buttons share the same `id="btnAddEdit"`. Differentiate rows via the `name` attribute: `input[name$='dgDadosTecnicos$ctl{idx}$btnAddEdit']`. Same for `ckIsNA` checkboxes.
- **Mídias `cmdFechar` closes the tab** — The close button has `onclick="window.close()"` which destroys the page immediately. Wrap the click in try/except.
- **Retry/recovery flow**: on failure, `process_item_with_retry` navigates home → worklist → retries. If that fails, it creates a fresh page.
- **Session expiration**: detected by `verificar_sessao()` polling for login indicators (including "Principal" and "Sair" links), pauses for manual re-login.
- **Timing instrumentation**: `StepTimer` in `main.py` logs per-step elapsed time, visual bar charts per item, running average, and ETA. Useful for identifying platform bottlenecks vs code bottlenecks.

## Configuration

All settings via `.env` file (see `.env.example`):
- `EXCEL_PATH` — input spreadsheet path
- `DOCUMENTS_DIR` — folder with documents to upload (use shared `downloads/` in production)
- `SHARED_DIR` — (future) shared network directory root
- `PROFILE_DIR` — Playwright session directory
- `SLOW_MO`, `HEADLESS`, `VIEWPORT_WIDTH`, `VIEWPORT_HEIGHT` — browser settings

## Known Issues

- **NCM validation**: Some NCM codes in the Excel are invalid/inactive in Klassmatt. The bot clears the field and continues, but the item will be incomplete. Items previously processed may have the NCM field as `readonly` — the bot detects and skips this.
- **Attribute table** (`dgDadosTecnicos`): Lives in `ITEM_Edita_DescricaoV3.aspx`, NOT in `ITEM_Edita.aspx`. Only visible after a PDM/pattern is set. The bot navigates to Descrições → Editar Descrição to access it, and navigates back via `butSIN_Voltar` → `Atuar no Item` after filling.
- **Attribute values not in tree**: Some Excel attribute values may not exist in the Klassmatt taxonomy tree. The bot logs a warning with available values and continues.
- **Reference autocomplete**: Company names in the Excel (e.g., "BAKER H") may not match the autocomplete suggestions in Klassmatt (e.g., "BAKER HUGHES"). This causes a timeout on the reference step.
