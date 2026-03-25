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

# Fix NCM on already-processed items (Retornar Etapa → NCM → Remeter)
python fix_ncm.py
```

First run requires manual login in the browser window; the session is saved to `./playwright_profile/`. On subsequent runs, the bot resumes from `progress.json`, skipping items with status "ok".

There are no automated tests.

## Architecture

**Async orchestrator pattern** — `main.py` drives a 12-step workflow per Excel row through Page Object modules.

### Core modules (root level)

- **main.py** — Entry point and orchestrator. `run()` loads Excel, validates documents, launches browser, iterates items with retry logic (3 attempts + backoff). `process_item()` calls page objects in sequence. Includes `StepTimer` class for per-step timing instrumentation. 5s delay between items to avoid overloading Klassmatt.
- **config.py** — All configuration: URLs, timeouts, Excel column mappings, and CSS selectors for ASP.NET elements. Also loads `.env`.
- **browser.py** — Playwright persistent context setup + helper functions (`safe_click` with JS fallback for overlay interception, `safe_fill`, `wait_for_text`, `retry_action`, `hide_overlays`, session verification, popup dismissal, dialog auto-accept).
- **excel_handler.py** — Reads Excel via openpyxl, auto-detects header row, maps columns including 30 dynamic attributes (Atrib_1_Valor..Atrib_30_Valor). Colors rows green/red/orange for success/error/skip. `validate_documents()` resolves doc paths against `DOCUMENTS_DIR`.
- **state.py** — JSON-based progress persistence (`progress.json`). Tracks item status: "ok", "error", "skipped", "duplicate".
- **logger.py** — Dual output: file at DEBUG (`klassmatt_rpa.log`), console at INFO.
- **fix_ncm.py** — Standalone script to fix NCM format on already-processed items. Handles Retornar Etapa → fill NCM → Remeter Modec flow.

### Page Objects (`pages/`)

Each module encapsulates interaction with a specific Klassmatt UI section:

| Module | Step |
|---|---|
| `item.py` | Search SIN via `#DIVResultado` grid, select via `abreSIN()` JS, create item (skips if already created), finalize & remit to MODEC (adapts flow to available buttons, gracefully handles missing buttons) |
| `classifications.py` | UNSPSC code via popup search/select (`#ckSelUNSPSC`) |
| `fiscal.py` | NCM field — formats `XXXXXXXX` → `XXXX.XX.XX`, validates after fill, clears if rejected by ASP.NET. Checks `is_editable()` to skip readonly fields. |
| `references.py` | Company reference + part number. Checks existing references before adding (idempotent). Uses `#iButAddRef` for new, `Imagebutton22` for edit. Multiple autocomplete fallbacks. |
| `relationships.py` | Old code relationship (Type=CÓDIGO ANTIGO, Status=ATIVO ERP, Comment=ZBRA). Retries tab navigation 3x. Detects existing relationships. |
| `media.py` | Document upload — detects new browser tab (`Midia.aspx`). Checks media count in tab label to skip if already uploaded (idempotent). |
| `descriptions.py` | SAP description validation (reads `#txtD2` for length check, "Exibe D2" toggle), PDM category change. All tab/button clicks via JS evaluate to bypass div1 overlay. |
| `attributes.py` | Up to 30 technical attributes — opens `Dt_EditaArvore.aspx` popup per attribute, navigates alphabet tree, selects value. All navigation via JS evaluate. |
| `worklist.py` | Navigate to Worklist, filter "Todas as Solicitações" via select2 + JS `pesquisar()` |

## Key Patterns

- **All automation is async/await** — Playwright async API throughout.
- **Selectors are centralized in `config.py`** — never hardcode selectors in page objects; add them to `config.SELECTORS`.
- **ASP.NET HTML IDs differ from `name` attributes** — The `name` uses hierarchy like `ctl00$Body$tabFiscal$txtNCMTIPI` but the rendered `id` is just `txtNCMTIPI`. Always use the actual `id` attribute for selectors, not the `name`. IDs are also case-sensitive (e.g., `Imagebutton7` not `imagebutton7`).
- **div1 overlay intercepts Playwright clicks** — `<div id="div1">` (description panel) sits on top of tabs and footer buttons. Use `hide_overlays(page)` after navigation, or `page.evaluate("el => el.click()")` as fallback. `safe_click()` in browser.py handles this automatically.
- **CRITICAL: hide_overlays BEFORE tab clicks, not just after** — Tab clicks via JS (`tab.click()`) are silently swallowed by the div1 overlay. The click appears to succeed (no error), but the tab content never loads. Always call `hide_overlays(page)` BEFORE clicking a tab, then verify the expected content actually appeared. If it didn't, retry with `__doPostBack` directly (e.g., `__doPostBack('ctl00$Body$dlTab$ctl01$lbutMenu', '')`). This was the root cause of PDM failures — `change_pdm()` clicked "Descrições" but the overlay ate the click, so "Editar Descrição" was never found.
- **Tab navigation: always verify, never assume** — After clicking an ASP.NET tab, check that the expected content loaded (e.g., a specific link or element). Tabs can fail silently due to overlays, dirty state alerts, or stale page context. Use `__doPostBack` as fallback — it bypasses overlays entirely. Known postback IDs: Descrições=`ctl00$Body$dlTab$ctl01$lbutMenu`, Editar Descrição=`ctl00$Body$tabDescricoes$lbutAlterarDescr`.
- **All steps must be idempotent** — Items may be reprocessed after errors. Each step checks existing state: reference count in tab label, media count, NCM `is_editable()`, relationship duplicate alert, dgDadosTecnicos presence for PDM.
- **NCM format: XXXX.XX.XX** — Excel has `73181500`, Klassmatt expects `7318.15.00`. `_format_ncm()` in fiscal.py converts automatically. Trigger validation via `getDescricaoNCM('NCM')`.
- **Worklist grid uses `<div>` not `<table>`** — Results are in `#DIVResultado > .result` divs, not `table.GridClass` rows. SIN links use `javascript:abreSIN({id})`.
- **Select2 dropdowns** — The worklist filter is a select2 widget wrapping a native `<select>`. `page.select_option()` changes the value but doesn't trigger the JS search; call `pesquisar(0, '')` via `page.evaluate()` after.
- **Mídias tab opens a new browser tab** — The link uses `OpenNewTab()` JS. Detect the new tab by checking `page.context.pages` for `Midia.aspx` in URL.
- **Reference buttons: `#iButAddRef` (ADD) vs `Imagebutton22` (EDIT)** — `iButAddRef` adds a new reference. `Imagebutton22` is inside `rptReferencias` repeater and edits existing references. The bot checks reference count before deciding.
- **Reference save may show warning page** — "Referência igual em fabricante/fornecedor/cliente diferente!" redirects to a page with Voltar/Continuar buttons. Use short timeout on `wait_for_load_state` and check for these buttons.
- **Items may already be created** — SINs in FINALIZACAO status already have an item. The bot detects this (no "Criar item" button) and skips the creation step.
- **Attribute tree popup (`Dt_EditaArvore.aspx`)** — The `dgDadosTecnicos` table lives in `ITEM_Edita_DescricaoV3.aspx` (not `ITEM_Edita.aspx`). Each attribute's edit button (`btnAddEdit`) calls `AbreJanTaxonomia()` which opens a popup via `window.open()`. The tree has a hierarchical structure: root node → alphabet letters (A-Z) → values. Clicking a letter triggers `__doPostBack` (full page reload in popup). Use JS `evaluate` to click the letter, then `wait_for_load_state("networkidle")` to wait for the postback. Use `evaluate` to find/click the value among potentially 1900+ nodes, then Playwright `click()` for the "Selecionar" button.
- **Attribute selectors use `name`, not `id`** — In `dgDadosTecnicos`, all `btnAddEdit` buttons share the same `id="btnAddEdit"`. Differentiate rows via the `name` attribute: `input[name$='dgDadosTecnicos$ctl{idx}$btnAddEdit']`. Same for `ckIsNA` checkboxes.
- **Mídias `cmdFechar` closes the tab** — The close button has `onclick="window.close()"` which destroys the page immediately. Wrap the click in try/except.
- **Retry/recovery flow**: on failure, `process_item_with_retry` navigates home → worklist → retries. If that fails, it creates a fresh page.
- **Session expiration**: detected by `verificar_sessao()` polling for login indicators (including "Principal" and "Sair" links), pauses for manual re-login.
- **Timing instrumentation**: `StepTimer` in `main.py` logs per-step elapsed time, visual bar charts per item, running average, and ETA. Useful for identifying platform bottlenecks vs code bottlenecks.
- **Klassmatt rate limiting** — Navigating too fast causes "Ocorreu uma exceção" errors. 5s delay between items. After "Remeter Modec", wait for `pesquisar()` function to be defined before searching next SIN.
- **Retornar Etapa** — Items in APROVACAO-TECNICA can be returned to FINALIZACAO via `#lkbutTrazerDeVolta`. Shows inline panel (not JS dialog) with "Sim"/"Não" buttons. Used by `fix_ncm.py`.

## Configuration

All settings via `.env` file (see `.env.example`):
- `EXCEL_PATH` — input spreadsheet path
- `DOCUMENTS_DIR` — folder with documents to upload (use shared `downloads/` in production)
- `SHARED_DIR` — (future) shared network directory root
- `PROFILE_DIR` — Playwright session directory
- `SLOW_MO`, `HEADLESS`, `VIEWPORT_WIDTH`, `VIEWPORT_HEIGHT` — browser settings (use 1920x1080 to avoid overlay issues)

## Known Issues

- **NCM format** (RESOLVED): Excel has NCM without dots (`73181500`). The bot now auto-formats to `XXXX.XX.XX` (`7318.15.00`). For items already processed, use `fix_ncm.py`.
- **Attribute table** (`dgDadosTecnicos`): Lives in `ITEM_Edita_DescricaoV3.aspx`, NOT in `ITEM_Edita.aspx`. Only visible after a PDM/pattern is set. The bot navigates to Descrições → Editar Descrição to access it, and navigates back via `butSIN_Voltar` → `Atuar no Item` after filling.
- **Attribute values not in tree**: Some Excel attribute values may not exist in the Klassmatt taxonomy tree. The bot logs a warning with available values and continues.
- **Reference autocomplete**: Company names in the Excel (e.g., "BAKER H") may not match the autocomplete suggestions in Klassmatt (e.g., "BAKER HUGHES"). The bot tries multiple fallbacks and continues without selection if none match — Klassmatt prompts to create new fabricante.
- **Unicode logging on Windows**: `StepTimer` bar charts use `█` and `→` characters that fail on cp1252 console. Cosmetic only — doesn't affect processing. Logs to file work fine (UTF-8).
- **Items in APROVACAO-TECNICA are skipped** — `process_item()` detects non-FINALIZACAO status and returns "skipped". `is_processed()` skips both "ok" and "skipped" items on subsequent runs. No longer attempts "Retornar Etapa".
- **KlassmattSessionError + browser restart** — `_voltar_worklist()` raises `KlassmattSessionError` on page errors or missing `pesquisar()`. `_restart_browser()` closes and reopens browser. `process_item_with_retry` propagates `(status, page, context, pw)` to support restart.
- **"Adicionar Mídia" timeout** — Most common permanent error (79% of errors). Link doesn't render after consecutive uploads on items with 5+ docs. Retry is 0% effective for this error.
- **Silent skip when "Empresa" is empty** — `main.py` line 174: `if item.get("empresa")` skips the reference step without logging a warning. 74 of 412 OK items had no reference filled because of this. Same items also had no relationship if `codigo_60` was empty.
- **NCM rejected by Klassmatt** — Some NCM codes (e.g., `84799090`, `73181500`, `84841000`) are rejected even after formatting. The bot clears the field and continues — item is marked OK but NCM is empty.
- **Attribute "dados técnicos" alert** — Klassmatt shows alert "É necessário preencher/verificar os dados técnicos destacados!" when attribute values don't exist in the taxonomy tree. Bot accepts the dialog and continues — item marked OK but attributes incomplete.
