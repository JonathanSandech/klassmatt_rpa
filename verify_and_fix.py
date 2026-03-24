"""Verifica e corrige SINs em uma única passagem.

Abre cada SIN, lê os dados (verify), e se divergente, corrige
apenas os campos errados (fix) — tudo na mesma sessão, sem
precisar navegar duas vezes.

Uso:
    python verify_and_fix.py                         # todos os SINs da planilha
    python verify_and_fix.py 474284 474291           # SINs específicos
    python verify_and_fix.py --file=lista.txt        # SINs de um arquivo
    python verify_and_fix.py --only-divergent        # re-processa divergentes do report
    python verify_and_fix.py --verify-only            # só verifica, não corrige
"""

import asyncio
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, Dialog

from config import (
    EXCEL_PATH, PROFILE_DIR, SLOW_MO, HEADLESS,
    VIEWPORT_WIDTH, VIEWPORT_HEIGHT, PROGRESS_FILE,
    RELATIONSHIP_TYPE,
)
from excel_handler import load_excel, validate_documents
from browser import hide_overlays, verificar_sessao
from logger import log

# Page objects (fix)
from pages.classifications import fill_unspsc
from pages.fiscal import fill_ncm
from pages.references import fill_reference
from pages.relationships import fill_relationship
from pages.media import upload_documents
from pages.descriptions import validate_sap_description, change_pdm
from pages.attributes import fill_attributes


REPORT_FILE = Path(__file__).parent / "verify_report.json"
REMETER_APOS_FIX = False


# ─── Report helpers ───────────────────────────────────────────


def _load_report() -> dict:
    if REPORT_FILE.exists():
        try:
            return json.loads(REPORT_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"timestamp": "", "total": 0, "counts": {}, "results": []}


def _save_report(report: dict):
    report["timestamp"] = datetime.now().isoformat()
    REPORT_FILE.write_text(
        json.dumps(report, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ─── Dialog ───────────────────────────────────────────────────


async def handle_dialog(dialog: Dialog):
    import browser as _browser
    _browser.last_dialog_message = dialog.message
    log.debug(f"Dialog ({dialog.type}): {dialog.message}")
    await dialog.accept()


# ─── Browser restart (mesmo padrão do main.py) ───────────────


async def _restart_browser(pw, context):
    """Fecha browser e reabre. Retorna (pw, context, page)."""
    log.warning("Fechando browser e reabrindo...")
    try:
        await context.close()
    except Exception:
        pass
    try:
        await pw.stop()
    except Exception:
        pass
    await asyncio.sleep(5)

    new_pw = await async_playwright().start()
    new_context = await new_pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=HEADLESS,
        slow_mo=SLOW_MO,
        viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
        args=[f"--window-size={VIEWPORT_WIDTH},{VIEWPORT_HEIGHT}"],
    )
    new_context.set_default_timeout(30_000)
    new_context.set_default_navigation_timeout(60_000)
    new_page = new_context.pages[0] if new_context.pages else await new_context.new_page()
    new_page.on("dialog", handle_dialog)

    await new_page.goto(
        "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
        wait_until="networkidle",
    )
    sessao_ok = await verificar_sessao(new_page)
    if not sessao_ok:
        log.warning("Sessão expirada após reabrir — aguardando 60s para re-login manual")
        await asyncio.sleep(60)
        await new_page.reload(wait_until="networkidle")

    await _navigate_to_worklist(new_page)
    # Selecionar "Todas as Solicitações"
    try:
        await new_page.select_option(
            "select:has(option[value='SOMENTE_REC_ACAO'])",
            label="Todas as Solicitações",
        )
        await new_page.evaluate("() => { pesquisar(0, ''); }")
        await new_page.wait_for_load_state("networkidle")
        await new_page.wait_for_timeout(3000)
    except Exception:
        pass

    return new_pw, new_context, new_page


# ─── Navigation helpers ───────────────────────────────────────


async def _check_page_error(page) -> bool:
    """Verifica se a página atual é a tela de erro do Klassmatt."""
    try:
        text = await page.evaluate("() => document.body.innerText.substring(0, 500)")
        return "exce" in text.lower() or "ACESSO" in text
    except Exception:
        return False


async def _navigate_to_worklist(page):
    """Garante que estamos na worklist com pesquisar() disponível."""
    for attempt in range(3):
        # Detectar página de erro do Klassmatt
        if await _check_page_error(page):
            log.warning("Página de erro detectada — voltando para Principal")
            await page.goto(
                "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
                wait_until="networkidle",
            )
            await page.wait_for_timeout(3000)

        ready = await page.evaluate("() => typeof pesquisar === 'function'")
        if ready:
            return
        await page.evaluate("""() => {
            const links = document.querySelectorAll('a');
            const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
            const principal = Array.from(links).find(a => a.innerText === 'Principal');
            if (wl) wl.click();
            else if (principal) principal.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)


async def _search_and_open_sin(page, sin: str):
    """Busca SIN na worklist e abre. Retorna a page do item (pode ser nova aba)."""
    await page.evaluate(f"""() => {{
        const ta = document.querySelector("textarea[name$='txtValor']");
        if (ta) ta.value = '{sin}';
        pesquisar(0, '');
    }}""")
    await page.wait_for_timeout(5000)

    pages_before = len(page.context.pages)
    found = await page.evaluate(f"""() => {{
        const link = document.querySelector("a[href*='abreSIN({sin})']");
        if (link) {{ link.click(); return true; }}
        const results = document.querySelectorAll('#DIVResultado .result a');
        if (results.length > 0) {{ results[0].click(); return true; }}
        return false;
    }}""")
    if not found:
        raise RuntimeError(f"SIN {sin} não encontrado na worklist")

    await page.wait_for_timeout(3000)

    # Retornar nova aba se abriu
    if len(page.context.pages) > pages_before:
        new_page = page.context.pages[-1]
        new_page.on("dialog", handle_dialog)
        await new_page.wait_for_load_state("networkidle")
        await new_page.wait_for_timeout(1000)
        return new_page

    await page.wait_for_load_state("networkidle")
    return page


async def _get_status(page) -> str:
    return await page.evaluate(
        "() => { const s = document.querySelector(\"input[id$='txtStatus']\"); return s ? s.value : ''; }"
    )


async def _voltar_worklist(page):
    """Volta para a worklist navegando via Principal."""
    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const p = Array.from(links).find(a => a.innerText.trim() === 'Principal');
        if (p) p.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const wl = Array.from(links).find(a => a.innerText.includes('Acompanhamento das Solicita'));
        if (wl) wl.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2000)

    try:
        await page.select_option(
            "select:has(option[value='SOMENTE_REC_ACAO'])",
            label="Todas as Solicitações",
        )
        await page.evaluate("() => { pesquisar(0, ''); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception:
        pass


async def _retornar_etapa(page) -> bool:
    """Retorna o item de APROVACAO-TECNICA para FINALIZACAO."""
    log.info("  Retornando etapa (APROVACAO-TECNICA → FINALIZACAO)...")
    await page.evaluate("""() => {
        const btn = document.querySelector('#lkbutTrazerDeVolta');
        if (btn) btn.click();
    }""")
    await page.wait_for_timeout(3000)

    sim_btn = page.locator("input[value='Sim']")
    try:
        await sim_btn.click(timeout=5000)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception:
        log.warning("  Botão 'Sim' não encontrado para Retornar Etapa")
        return False

    status = await _get_status(page)
    log.info(f"  Status após retorno: {status}")
    return status == "FINALIZACAO"


# ─── Verify: leitura de campos ────────────────────────────────


def _format_ncm(ncm_raw: str) -> str:
    digits = re.sub(r"\D", "", str(ncm_raw))
    if len(digits) == 8:
        return f"{digits[:4]}.{digits[4:6]}.{digits[6:8]}"
    return str(ncm_raw)


async def _read_unspsc(page) -> str:
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Classificações'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
        const inp = document.querySelector('#txtUNSPSC');
        if (inp && inp.value) {
            const m = inp.value.match(/^(\\d{8})/);
            if (m) return m[1];
        }
        return '';
    }""")


async def _read_ncm(page) -> str:
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Fiscal'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
        const ncm = document.querySelector('#txtNCMTIPI');
        return ncm ? ncm.value.trim() : '';
    }""")


async def _read_references(page) -> list[dict]:
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Referências'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
        const result = [];
        const allText = document.body.innerText || '';
        const matches = allText.match(/Referência\\/Fabricante:\\s*([^\\n]+)/g);
        if (matches) {
            for (const m of matches) {
                const val = m.replace('Referência/Fabricante:', '').trim();
                if (val === 'N/A' || val === 'N/A/N/A' || val === '/' || val === '') continue;
                const lastSlash = val.lastIndexOf('/');
                let partNumber, empresa;
                if (lastSlash > 0) {
                    partNumber = val.substring(0, lastSlash).trim();
                    empresa = val.substring(lastSlash + 1).trim();
                } else {
                    partNumber = val;
                    empresa = '';
                }
                result.push({partNumber, empresa, raw: val});
            }
        }
        return result;
    }""")


async def _read_relationships(page) -> list[dict]:
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Relacionamentos'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
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
            if (!tipo && !codigo) continue;
            result.push({tipo, codigo, status, comentario});
        }
        return result;
    }""")


async def _read_media_count(page) -> int:
    return await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        for (const tab of tabs) {
            const m = tab.innerText.match(/Mídias\\s*\\((\\d+)\\)/);
            if (m) return parseInt(m[1]);
        }
        return -1;
    }""")


# ─── Core: verify + fix em uma passagem ───────────────────────


async def verify_and_fix_sin(
    page, sin: str, item: dict, *, verify_only: bool = False,
) -> dict:
    """Verifica um SIN. Se divergente e não verify_only, corrige na hora."""
    result = {
        "sin": sin,
        "status": "ok",
        "item_status": "",
        "diffs": [],
        "fixed": [],
        "warnings": [],
        "elapsed": 0,
    }
    start = time.time()
    worklist_page = page
    item_page = page

    try:
        await _navigate_to_worklist(page)

        # Abrir SIN
        item_page = await _search_and_open_sin(page, sin)
        log.debug(f"  Aba do item: {item_page.url}")

        # Status
        item_status = await _get_status(item_page)
        result["item_status"] = item_status

        # Atuar no Item via __doPostBack (btn.click() não trigga postback ASP.NET)
        await item_page.evaluate("""() => {
            window.confirm = () => true;
            window.alert = () => {};
            __doPostBack('ctl00$Body$butAcao3', '');
        }""")
        await item_page.wait_for_load_state("networkidle")
        await item_page.wait_for_timeout(2000)

        # Detectar se Atuar abriu outra aba
        for p in page.context.pages:
            if p != worklist_page and p != item_page and "ITEM_Edita" in p.url:
                item_page = p
                await item_page.wait_for_load_state("networkidle")
                break

        await hide_overlays(item_page)
        log.debug(f"  Página de edição: {item_page.url}")

        # ── VERIFY: ler todos os campos ──

        # UNSPSC
        needs_unspsc = False
        if item.get("unspsc"):
            actual_unspsc = await _read_unspsc(item_page)
            expected_unspsc = str(item["unspsc"]).strip()
            if expected_unspsc.upper() != actual_unspsc.upper():
                result["diffs"].append({
                    "field": "UNSPSC", "expected": expected_unspsc, "actual": actual_unspsc,
                })
                needs_unspsc = True

        # NCM
        needs_ncm = False
        if item.get("ncm"):
            actual_ncm = await _read_ncm(item_page)
            expected_ncm = _format_ncm(str(item["ncm"]))
            if expected_ncm.upper() != actual_ncm.upper():
                result["diffs"].append({
                    "field": "NCM", "expected": expected_ncm, "actual": actual_ncm,
                })
                needs_ncm = True

        # Referências
        needs_reference = False
        actual_refs = await _read_references(item_page)
        if item.get("part_number"):
            expected_pn = str(item["part_number"]).strip()
            found_pn = any(
                expected_pn in ref.get("partNumber", "") or expected_pn in ref.get("raw", "")
                for ref in actual_refs
            )
            if not found_pn:
                result["diffs"].append({
                    "field": "Referência",
                    "expected": f"{item.get('empresa', '')} / {expected_pn}",
                    "actual": json.dumps(actual_refs, ensure_ascii=False) if actual_refs else "(nenhuma)",
                })
                needs_reference = True

        # Relacionamentos
        needs_relationship = False
        actual_rels = await _read_relationships(item_page)
        if item.get("codigo_60"):
            expected_code = str(item["codigo_60"]).strip()
            matching = [r for r in actual_rels if r["tipo"].upper() == RELATIONSHIP_TYPE.upper()]
            if not matching or not any(r["codigo"] == expected_code for r in matching):
                result["diffs"].append({
                    "field": "Relacionamento",
                    "expected": f"{RELATIONSHIP_TYPE} / {expected_code}",
                    "actual": ", ".join(f"{r['tipo']}/{r['codigo']}" for r in matching) or "(nenhum)",
                })
                needs_relationship = True

        # Mídias
        needs_media = False
        actual_media = await _read_media_count(item_page)
        doc_files = item.get("_doc_files", [])
        expected_docs = len(doc_files) if doc_files else 0
        if expected_docs > 0 and actual_media < expected_docs:
            result["diffs"].append({
                "field": "Mídias", "expected": str(expected_docs), "actual": str(actual_media),
            })
            needs_media = True
        elif actual_media > expected_docs and expected_docs > 0:
            result["warnings"].append(
                f"Mídias a mais: esperado {expected_docs}, encontrado {actual_media}"
            )

        # PDM / Atributos — verificação adiada para fase 3 (dentro de DescricaoV3)
        # Aqui só marcamos que PODE precisar verificar PDM/atributos
        has_pdm_in_excel = bool(item.get("pdm"))
        has_attrs_in_excel = bool(item.get("attributes"))
        needs_pdm = False
        needs_attributes = False

        # ── Se não há diffs non-PDM E não tem PDM na planilha → OK ──
        if not result["diffs"] and not has_pdm_in_excel and not has_attrs_in_excel:
            result["status"] = "ok"
            result["elapsed"] = round(time.time() - start, 1)
            return result

        # ── Se verify_only e não tem PDM para checar, encerrar ──
        if verify_only and not has_pdm_in_excel and not has_attrs_in_excel:
            result["status"] = "divergente" if result["diffs"] else "ok"
            result["elapsed"] = round(time.time() - start, 1)
            return result

        # ── FIX (ou verify_only com PDM para checar): ──
        has_non_pdm_diffs = bool(result["diffs"])
        if has_non_pdm_diffs and not verify_only:
            log.info(f"  Divergências: {', '.join(d['field'] for d in result['diffs'])} — corrigindo...")

        # Checar se precisa retornar etapa
        if item_status == "APROVACAO-TECNICA" and not verify_only:
            # Voltar para SIN_Item_Resultante para retornar etapa
            await item_page.evaluate("""() => {
                const links = document.querySelectorAll('a');
                const v = Array.from(links).find(a => a.innerText.trim() === 'Voltar');
                if (v) v.click();
            }""")
            await item_page.wait_for_load_state("networkidle")
            await item_page.wait_for_timeout(2000)

            ok = await _retornar_etapa(item_page)
            if not ok:
                result["status"] = "fix_failed"
                result["elapsed"] = round(time.time() - start, 1)
                return result

            # Re-entrar em edição
            await item_page.evaluate("""() => {
                window.confirm = () => true;
                window.alert = () => {};
                __doPostBack('ctl00$Body$butAcao3', '');
            }""")
            await item_page.wait_for_load_state("networkidle")
            await item_page.wait_for_timeout(2000)
            await hide_overlays(item_page)

        # Corrigir cada campo divergente (non-PDM)
        try:
            if not verify_only:
                if needs_unspsc and item.get("unspsc"):
                    await fill_unspsc(item_page, str(item["unspsc"]))
                    result["fixed"].append("UNSPSC")
                    log.info(f"  ✓ UNSPSC corrigido: {item['unspsc']}")

                if needs_ncm and item.get("ncm"):
                    await fill_ncm(item_page, str(item["ncm"]))
                    result["fixed"].append("NCM")
                    log.info(f"  ✓ NCM corrigido: {item['ncm']}")

                if needs_reference and item.get("empresa") and item.get("part_number"):
                    await fill_reference(item_page, str(item["empresa"]), str(item["part_number"]))
                    result["fixed"].append("Referência")
                    log.info(f"  ✓ Referência corrigida: {item['empresa']} / {item['part_number']}")

                if needs_relationship and item.get("codigo_60"):
                    await fill_relationship(item_page, str(item["codigo_60"]))
                    result["fixed"].append("Relacionamento")
                    log.info(f"  ✓ Relacionamento corrigido: {item['codigo_60']}")

                if needs_media and doc_files:
                    await upload_documents(item_page, doc_files)
                    result["fixed"].append("Mídias")
                    log.info(f"  ✓ Mídias: {len(doc_files)} doc(s)")

            # Salvar geral para limpar dirty state antes de navegar para PDM/Descrições
            if (needs_reference or needs_relationship) and (has_pdm_in_excel or has_attrs_in_excel):
                log.debug("  Salvando item para limpar dirty state...")
                for save_attempt in range(3):
                    try:
                        # Override confirm/alert ANTES de salvar (postback pode triggar alerts)
                        await item_page.evaluate("""() => {
                            window.confirm = () => true;
                            window.alert = () => {};
                        }""")
                        salvar_footer = item_page.locator("#butSalvar")
                        if await salvar_footer.count() > 0:
                            await salvar_footer.click()
                            await item_page.wait_for_load_state("networkidle")
                            await item_page.wait_for_timeout(2000)
                            # Re-override após postback (o JS é recarregado)
                            await item_page.evaluate("""() => {
                                window.confirm = () => true;
                                window.alert = () => {};
                            }""")
                            await hide_overlays(item_page)
                    except Exception as save_err:
                        log.warning(f"  Salvar geral falhou: {save_err}")

                    # Verificar se dirty state foi limpo tentando navegar para Classificações
                    dirty = await item_page.evaluate("""() => {
                        // Tentar clicar numa aba neutra para testar se há dirty state
                        // Se window.__doPostBack está bloqueado por alert, retorna true
                        try {
                            const tabs = document.querySelectorAll('a');
                            const tab = Array.from(tabs).find(a => a.innerText.includes('Classificações'));
                            if (tab) { tab.click(); return false; }
                        } catch(e) { return true; }
                        return false;
                    }""")
                    await item_page.wait_for_timeout(1500)

                    # Checar se alert de dirty state apareceu (o dialog handler aceita, mas a página não navega)
                    still_dirty = await item_page.evaluate("""() => {
                        // Se ainda estamos na mesma aba (referências visível), dirty state persiste
                        const refTab = document.querySelector('#tabReferencias');
                        if (refTab && refTab.offsetHeight > 0) return true;
                        return false;
                    }""")

                    if not still_dirty:
                        log.debug(f"  Item salvo (dirty state limpo, tentativa {save_attempt + 1})")
                        break
                    else:
                        log.warning(f"  Dirty state persiste após save (tentativa {save_attempt + 1})")
                        # Forçar: navegar para aba de referências e salvar lá
                        await item_page.evaluate("""() => {
                            window.confirm = () => true;
                            window.alert = () => {};
                            // Clicar na aba referências para voltar ao contexto
                            const tabs = document.querySelectorAll('a');
                            const tab = Array.from(tabs).find(a => a.innerText.includes('Referências'));
                            if (tab) tab.click();
                        }""")
                        await item_page.wait_for_timeout(2000)

            # ── FASE 3: Verify + Fix PDM/Atributos (entrada única em DescricaoV3) ──
            if has_pdm_in_excel or has_attrs_in_excel:
                # Navegar para DescricaoV3
                await item_page.evaluate("""() => {
                    window.confirm = () => true;
                    window.alert = () => {};
                    const tabs = document.querySelectorAll('a');
                    const tab = Array.from(tabs).find(a => a.innerText.includes('Descrições'));
                    if (tab) tab.click();
                }""")
                await item_page.wait_for_load_state("networkidle")
                await item_page.wait_for_timeout(1000)

                if "ITEM_Edita_DescricaoV3" not in item_page.url:
                    await item_page.evaluate("""() => {
                        const links = document.querySelectorAll('a');
                        const link = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
                        if (link) link.click();
                    }""")
                    await item_page.wait_for_load_state("networkidle")
                    await item_page.wait_for_timeout(1000)

                # Ler PDM e atributos in-place
                pdm_data = await item_page.evaluate("""() => {
                    const data = {padronizado: true, attributes: []};
                    if (document.body.innerText.includes('NÃO-PADRONIZADO'))
                        data.padronizado = false;
                    const dg = document.querySelector('#dgDadosTecnicos');
                    if (!dg) return data;
                    const rows = dg.querySelectorAll('tr');
                    for (let i = 1; i < rows.length; i++) {
                        const cells = rows[i].querySelectorAll('td');
                        if (cells.length < 2) continue;
                        // Label pode estar em cells[0] ou cells[1] (cells[0] pode ser ícone LED)
                        let label = (cells[0]?.innerText || '').trim();
                        if (!label || label.length < 2)
                            label = (cells[1]?.innerText || '').trim();
                        if (!label || label === 'Dados Técnicos') continue;
                        const idx = (i + 1).toString().padStart(2, '0');
                        const hidden = document.querySelector(
                            `input[name$='dgDadosTecnicos$ctl${idx}$hdnDtTexto']`
                        );
                        const naCheckbox = document.querySelector(
                            `input[name$='dgDadosTecnicos$ctl${idx}$ckIsNA']`
                        );
                        const value = hidden ? hidden.value.trim() : '';
                        const isNA = naCheckbox ? naCheckbox.checked : false;
                        data.attributes.push({label, value: value || (isNA ? 'N/A' : ''), isNA});
                    }
                    return data;
                }""")

                # Computar diffs de PDM
                if has_pdm_in_excel and not pdm_data.get("padronizado", True):
                    result["diffs"].append({
                        "field": "PDM",
                        "expected": str(item["pdm"]),
                        "actual": "(NÃO-PADRONIZADO)",
                    })
                    needs_pdm = True
                    needs_attributes = True  # PDM errado → atributos precisam ser refeitos

                # Computar diffs de atributos
                if not needs_pdm:
                    expected_attrs = item.get("attributes", [])
                    actual_attrs = pdm_data.get("attributes", [])
                    for i, attr in enumerate(actual_attrs):
                        if i >= len(expected_attrs):
                            break
                        exp_val = expected_attrs[i]
                        if exp_val is None or (isinstance(exp_val, str) and exp_val.strip() == ""):
                            exp_val = "N/A"
                        else:
                            exp_val = str(exp_val).strip()
                        act_val = attr.get("value", "") or ("N/A" if attr.get("isNA") else "")
                        if exp_val.upper() != act_val.upper():
                            result["diffs"].append({
                                "field": f"Atributo: {attr.get('label', f'Atrib_{i+1}')}",
                                "expected": exp_val, "actual": act_val or "(vazio)",
                            })
                            needs_attributes = True

                # Log diffs de PDM/atributos
                if needs_pdm or needs_attributes:
                    pdm_diffs = [d["field"] for d in result["diffs"] if "PDM" in d["field"] or "Atributo" in d["field"]]
                    if not has_non_pdm_diffs:
                        log.info(f"  Divergências: {', '.join(pdm_diffs)} — {'corrigindo...' if not verify_only else 'verify only'}")
                    else:
                        log.info(f"  + PDM/Atributos: {', '.join(pdm_diffs)}")

                # Se verify_only, não corrigir
                if verify_only:
                    pass  # diffs já registrados, status final será computado abaixo
                elif needs_pdm and item.get("pdm"):
                    # Já estamos em DescricaoV3 — change_pdm detecta e pula navegação
                    pdm_ok = await change_pdm(item_page, str(item["pdm"]))
                    if pdm_ok:
                        result["fixed"].append("PDM")
                        log.info(f"  ✓ PDM corrigido: {item['pdm']}")
                    else:
                        result["warnings"].append("PDM falhou")
                        log.warning(f"  ✗ PDM NÃO corrigido: {item['pdm']}")
                        needs_attributes = False  # Sem PDM, não preencher atributos

                if needs_attributes and not verify_only:
                    # Já estamos em DescricaoV3 — fill_attributes detecta e pula navegação
                    attr_ok = await fill_attributes(item_page, item.get("attributes", []))
                    if attr_ok:
                        result["fixed"].append("Atributos")
                        log.info("  ✓ Atributos corrigidos")
                    else:
                        result["warnings"].append("Atributos falhou")
                        log.warning("  ✗ Atributos NÃO corrigidos")

            # Remeter
            if REMETER_APOS_FIX:
                remeter = item_page.locator("input[value='Remeter Modec']")
                if await remeter.count() > 0:
                    await remeter.click()
                    await item_page.wait_for_timeout(3000)
                    sim_btn = item_page.locator("input[value='Sim']")
                    try:
                        await sim_btn.click(timeout=5000)
                        await item_page.wait_for_load_state("networkidle")
                        result["fixed"].append("Remetido")
                    except Exception:
                        pass

            # Status final
            if verify_only:
                result["status"] = "divergente" if result["diffs"] else "ok"
            elif result["warnings"]:
                result["status"] = "parcial"
            elif result["fixed"]:
                result["status"] = "corrigido"
            elif not result["diffs"]:
                result["status"] = "ok"
            else:
                result["status"] = "divergente"

        except Exception as fix_err:
            log.error(f"  Erro ao corrigir: {fix_err}")
            result["status"] = "fix_failed"
            result["warnings"].append(f"Fix falhou: {str(fix_err)[:200]}")

    except Exception as e:
        result["status"] = "error"
        result["diffs"].append({"field": "ERRO", "expected": "", "actual": str(e)})
        log.error(f"  Erro ao processar SIN {sin}: {e}")

    finally:
        # Fechar abas extras
        if item_page != worklist_page:
            try:
                if not item_page.is_closed():
                    await item_page.close()
            except Exception:
                pass

    result["elapsed"] = round(time.time() - start, 1)
    return result


# ─── Main ─────────────────────────────────────────────────────


async def run():
    log.info("=" * 60)
    log.info("Verify & Fix — Início")
    log.info("=" * 60)

    verify_only = "--verify-only" in sys.argv
    only_divergent = "--only-divergent" in sys.argv
    cli_sins = [s for s in sys.argv[1:] if not s.startswith("--")]

    # Suporte a --file=lista.txt
    for arg in sys.argv[1:]:
        if arg.startswith("--file="):
            filepath = Path(arg.split("=", 1)[1])
            cli_sins = [s.strip() for s in filepath.read_text().splitlines() if s.strip()]

    # Ler Excel
    wb, items = load_excel()
    items = validate_documents(items)
    sin_data = {}
    for item in items:
        sin = str(item.get("sin", ""))
        if sin:
            sin_data[sin] = item

    # Carregar report existente
    report = _load_report()

    if only_divergent:
        divergent_sins = [
            r["sin"] for r in report.get("results", [])
            if r.get("status") in ("divergente", "error", "fix_failed")
        ]
        # Remover do report os que vamos re-processar
        report["results"] = [
            r for r in report.get("results", [])
            if r.get("status") not in ("divergente", "error", "fix_failed")
        ]
        sins_to_process = [s for s in divergent_sins if s in sin_data]
        log.info(f"--only-divergent: {len(sins_to_process)} SINs a re-processar")
    elif cli_sins:
        sins_to_process = cli_sins
        # Remover resultados antigos desses SINs
        sins_set = set(sins_to_process)
        report["results"] = [r for r in report.get("results", []) if r["sin"] not in sins_set]
        log.info(f"{len(sins_to_process)} SINs via CLI/arquivo")
    else:
        sins_to_process = list(sin_data.keys())

    if verify_only:
        log.info("Modo: VERIFY ONLY (sem correção)")
    else:
        log.info("Modo: VERIFY + FIX")

    log.info(f"SINs a processar: {len(sins_to_process)}")

    # Recontabilizar counts
    counts = {}
    for r in report.get("results", []):
        s = r.get("status", "error")
        counts[s] = counts.get(s, 0) + 1

    # Browser
    pw = await async_playwright().start()
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=HEADLESS,
        slow_mo=SLOW_MO,
        viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
        args=[f"--window-size={VIEWPORT_WIDTH},{VIEWPORT_HEIGHT}"],
    )
    context.set_default_timeout(30_000)
    context.set_default_navigation_timeout(60_000)
    page = context.pages[0] if context.pages else await context.new_page()
    page.on("dialog", handle_dialog)

    # Navegar para worklist
    try:
        await page.goto(
            "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
            wait_until="networkidle",
        )
        await page.click("text=Acompanhamento das Solicitações (Worklist)")
        await page.wait_for_load_state("networkidle")
        await page.select_option(
            "select:has(option[value='SOMENTE_REC_ACAO'])",
            label="Todas as Solicitações",
        )
        await page.evaluate("() => { pesquisar(0, ''); }")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(3000)
    except Exception as e:
        log.error(f"Falha ao inicializar worklist: {e}")
        log.info("Aguardando 60s para login manual...")
        await asyncio.sleep(60)

    all_results = report.get("results", [])

    for i, sin in enumerate(sins_to_process):
        if sin not in sin_data:
            log.warning(f"SIN {sin} não encontrado na planilha — pulando")
            counts["not_found"] = counts.get("not_found", 0) + 1
            continue

        # Verificar sessão a cada 20 itens
        if i % 20 == 0:
            sessao_ok = await verificar_sessao(page)
            if not sessao_ok:
                log.warning("Sessão expirada — aguardando 60s para login manual...")
                await asyncio.sleep(60)
                try:
                    await page.reload(wait_until="networkidle")
                    sessao_ok = await verificar_sessao(page)
                except Exception:
                    sessao_ok = False
                if not sessao_ok:
                    log.error("Sessão não recuperada após 60s — encerrando")
                    break

        log.info(f"[{i+1}/{len(sins_to_process)}] SIN {sin}...")

        try:
            result = await verify_and_fix_sin(
                page, sin, sin_data[sin], verify_only=verify_only,
            )
            all_results.append(result)
            counts[result["status"]] = counts.get(result["status"], 0) + 1

            # Log
            st = result["status"]
            elapsed = result["elapsed"]
            if st == "ok":
                log.info(f"  ✓ OK ({elapsed:.1f}s)")
            elif st == "corrigido":
                fixed_str = ", ".join(result["fixed"])
                log.info(f"  ✓ CORRIGIDO: {fixed_str} ({elapsed:.1f}s)")
            elif st == "divergente":
                diffs_str = ", ".join(d["field"] for d in result["diffs"])
                log.warning(f"  ✗ DIVERGENTE: {diffs_str} ({elapsed:.1f}s)")
            elif st == "fix_failed":
                diffs_str = ", ".join(d["field"] for d in result["diffs"])
                log.error(f"  ✗ FIX FALHOU: {diffs_str} ({elapsed:.1f}s)")
            else:
                log.error(f"  ! ERRO ({elapsed:.1f}s)")

            # Voltar para worklist
            try:
                await _voltar_worklist(page)
                # Checar se caiu em página de erro após voltar
                if await _check_page_error(page):
                    raise Exception("Página de erro após voltar para worklist")
            except Exception as nav_err:
                log.warning(f"Erro ao voltar para worklist: {nav_err} — reiniciando browser")
                try:
                    pw, context, page = await _restart_browser(pw, context)
                except Exception:
                    log.error("Restart browser falhou — encerrando")
                    break
            await asyncio.sleep(5)

        except Exception as e:
            log.error(f"Erro fatal SIN {sin}: {e}")
            counts["error"] = counts.get("error", 0) + 1
            all_results.append({
                "sin": sin, "status": "error",
                "diffs": [{"field": "ERRO FATAL", "expected": "", "actual": str(e)}],
                "fixed": [], "warnings": [], "elapsed": 0,
            })
            try:
                pw, context, page = await _restart_browser(pw, context)
            except Exception:
                log.error("Restart browser falhou — encerrando")
                break

        # Salvar incrementalmente
        report["total"] = len(sins_to_process)
        report["counts"] = counts
        report["results"] = all_results
        _save_report(report)

    # Resumo
    log.info("=" * 60)
    log.info("RESUMO")
    log.info("=" * 60)
    log.info(f"  Total processados: {sum(counts.values())}")
    for status_name in ("ok", "corrigido", "divergente", "fix_failed", "error"):
        if counts.get(status_name, 0) > 0:
            log.info(f"  {status_name:15s}: {counts[status_name]}")
    log.info("=" * 60)

    await context.close()
    await pw.stop()


if __name__ == "__main__":
    asyncio.run(run())
