"""Verificação cara-crachá: compara dados da planilha com o Klassmatt.

Abre cada SIN no sistema, lê os dados preenchidos via JS evaluate,
e compara com os valores esperados da planilha. Gera relatório de
discrepâncias sem alterar nenhum dado.

Uso:
    python verify_items.py                     # verifica todos os SINs da planilha
    python verify_items.py 474284 474291       # verifica SINs específicos
    python verify_items.py --from-progress     # verifica SINs com status "ok" no progress.json
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
from excel_handler import load_excel
from logger import log


REPORT_FILE = Path(__file__).parent / "verify_report.json"


async def handle_dialog(dialog: Dialog):
    log.debug(f"Dialog ({dialog.type}): {dialog.message}")
    await dialog.accept()


async def _navigate_to_worklist(page):
    """Garante que estamos na worklist com pesquisar() disponível."""
    for attempt in range(3):
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
    """Busca SIN na worklist e abre a página do item."""
    await page.evaluate(f"""() => {{
        const ta = document.querySelector("textarea[name$='txtValor']");
        if (ta) ta.value = '{sin}';
        pesquisar(0, '');
    }}""")
    await page.wait_for_timeout(5000)

    found = await page.evaluate(f"""() => {{
        const link = document.querySelector("a[href*='abreSIN({sin})']");
        if (link) {{ link.click(); return true; }}
        const results = document.querySelectorAll('#DIVResultado .result a');
        if (results.length > 0) {{ results[0].click(); return true; }}
        return false;
    }}""")
    if not found:
        raise RuntimeError(f"SIN {sin} não encontrado na worklist")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(3000)


async def _read_status(page) -> str:
    """Lê status do item."""
    return await page.evaluate(
        "() => { const s = document.querySelector(\"input[id$='txtStatus']\"); return s ? s.value : ''; }"
    )


async def _read_unspsc(page) -> str:
    """Lê UNSPSC da aba Classificações."""
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Classificações'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
        // Input #txtUNSPSC contém "31401503. O ring molded gasket" — extrair 8 dígitos
        const inp = document.querySelector('#txtUNSPSC');
        if (inp && inp.value) {
            const m = inp.value.match(/^(\\d{8})/);
            if (m) return m[1];
        }
        // Fallback: procurar em qualquer input/td/span com 8 dígitos
        const cells = document.querySelectorAll('td, span, input');
        for (const cell of cells) {
            const text = (cell.value || cell.innerText || '').trim();
            const m = text.match(/^(\\d{8})/);
            if (m) return m[1];
        }
        return '';
    }""")


async def _read_ncm(page) -> str:
    """Lê NCM da aba Fiscal."""
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
    """Lê referências existentes."""
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Referências'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    return await page.evaluate("""() => {
        const result = [];
        // divReferencias contém "Referência/Fabricante: {partNumber}/{empresa}"
        const div = document.querySelector('#divReferencias');
        if (div) {
            const text = div.innerText || '';
            // Pode ter múltiplas referências; cada uma tem "Referência/Fabricante:"
            const matches = text.match(/Referência\\/Fabricante:\\s*([^\\n]+)/g);
            if (matches) {
                for (const m of matches) {
                    const val = m.replace('Referência/Fabricante:', '').trim();
                    const parts = val.split('/');
                    const partNumber = (parts[0] || '').trim();
                    const empresa = parts.slice(1).join('/').trim();
                    result.push({partNumber, empresa, raw: val});
                }
            }
        }
        return result;
    }""")


async def _read_relationships(page) -> list[dict]:
    """Lê relacionamentos existentes."""
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
    """Lê quantidade de mídias pelo label da aba."""
    count = await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        for (const tab of tabs) {
            const text = tab.innerText;
            const m = text.match(/Mídias\\s*\\((\\d+)\\)/);
            if (m) return parseInt(m[1]);
        }
        return -1;
    }""")
    return count


async def _read_pdm_and_attributes(page) -> dict:
    """Lê PDM e atributos da aba Descrições."""
    # Navegar para Descrições
    await page.evaluate("""() => {
        const tabs = document.querySelectorAll('a');
        const tab = Array.from(tabs).find(a => a.innerText.includes('Descrições'));
        if (tab) tab.click();
    }""")
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    # Tentar entrar na DescricaoV3
    found = await page.evaluate("""() => {
        const links = document.querySelectorAll('a');
        const link = Array.from(links).find(a => a.innerText.includes('Editar Descri'));
        if (link) { link.click(); return true; }
        return false;
    }""")
    if found:
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1000)

    result = await page.evaluate("""() => {
        const data = {padronizado: true, attributes: []};

        // Verificar se é NÃO-PADRONIZADO
        if (document.body.innerText.includes('NÃO-PADRONIZADO')) {
            data.padronizado = false;
        }

        // Ler atributos do dgDadosTecnicos
        const dg = document.querySelector('#dgDadosTecnicos');
        if (!dg) return data;

        const rows = dg.querySelectorAll('tr');
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].querySelectorAll('td');
            if (cells.length < 2) continue;
            const label = (cells[0]?.innerText || '').trim();
            if (!label || label === 'Dados Técnicos') continue;

            // Ler valor do hidden field
            const idx = (i + 1).toString().padStart(2, '0');
            const hidden = document.querySelector(
                `input[name$='dgDadosTecnicos$ctl${idx}$hdnDtTexto']`
            );
            const naCheckbox = document.querySelector(
                `input[name$='dgDadosTecnicos$ctl${idx}$ckIsNA']`
            );
            const value = hidden ? hidden.value.trim() : '';
            const isNA = naCheckbox ? naCheckbox.checked : false;

            data.attributes.push({
                label: label,
                value: value || (isNA ? 'N/A' : ''),
                isNA: isNA,
            });
        }
        return data;
    }""")

    # Voltar da DescricaoV3 se entramos
    if found and "ITEM_Edita_DescricaoV3" in page.url:
        await page.evaluate("""() => {
            const btn = document.querySelector('#butSIN_Voltar');
            if (btn) btn.click();
        }""")
        try:
            await page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass
        await page.wait_for_timeout(1000)
        # Atuar no Item se necessário
        await page.evaluate("""() => {
            const btn = document.querySelector("input[value='Atuar no Item']");
            if (btn) btn.click();
        }""")
        try:
            await page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass

    return result


def _format_ncm(ncm_raw: str) -> str:
    """Formata NCM de XXXXXXXX para XXXX.XX.XX."""
    digits = re.sub(r"\D", "", str(ncm_raw))
    if len(digits) == 8:
        return f"{digits[:4]}.{digits[4:6]}.{digits[6:8]}"
    return str(ncm_raw)


def _compare(expected, actual, field: str) -> dict | None:
    """Compara valor esperado vs real. Retorna diff dict ou None se ok."""
    exp = str(expected).strip() if expected else ""
    act = str(actual).strip() if actual else ""
    if exp and exp.upper() != act.upper():
        return {"field": field, "expected": exp, "actual": act}
    return None


async def verify_sin(page, sin: str, item: dict) -> dict:
    """Verifica um SIN no Klassmatt vs planilha. Retorna dict com resultado."""
    result = {
        "sin": sin,
        "status": "ok",
        "item_status": "",
        "diffs": [],
        "warnings": [],
        "elapsed": 0,
    }
    start = time.time()

    try:
        await _navigate_to_worklist(page)
        await _search_and_open_sin(page, sin)

        # Status
        item_status = await _read_status(page)
        result["item_status"] = item_status

        # Atuar no Item para acessar os dados
        await page.evaluate("""() => {
            const btn = document.querySelector("input[value='Atuar no Item']");
            if (btn) btn.click();
        }""")
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2000)

        # Esconder overlays
        await page.evaluate("""() => {
            const div1 = document.querySelector('#div1');
            if (div1) div1.style.pointerEvents = 'none';
            const pg2 = document.querySelector('#pg-2');
            if (pg2) pg2.style.pointerEvents = 'none';
        }""")

        # ── UNSPSC ──
        if item.get("unspsc"):
            actual_unspsc = await _read_unspsc(page)
            expected_unspsc = str(item["unspsc"]).strip()
            diff = _compare(expected_unspsc, actual_unspsc, "UNSPSC")
            if diff:
                result["diffs"].append(diff)

        # ── NCM ──
        if item.get("ncm"):
            actual_ncm = await _read_ncm(page)
            expected_ncm = _format_ncm(str(item["ncm"]))
            diff = _compare(expected_ncm, actual_ncm, "NCM")
            if diff:
                result["diffs"].append(diff)

        # ── Referências ──
        actual_refs = await _read_references(page)
        if item.get("part_number"):
            expected_pn = str(item["part_number"]).strip()
            found_pn = any(
                expected_pn in ref.get("partNumber", "") or expected_pn in ref.get("raw", "")
                for ref in actual_refs
            )
            if not found_pn:
                result["diffs"].append({
                    "field": "Referência (Part Number)",
                    "expected": expected_pn,
                    "actual": json.dumps(actual_refs, ensure_ascii=False) if actual_refs else "(nenhuma)",
                })
        if not actual_refs and item.get("empresa"):
            result["diffs"].append({
                "field": "Referência",
                "expected": f"{item.get('empresa')} / {item.get('part_number')}",
                "actual": "(nenhuma referência)",
            })

        # ── Relacionamentos ──
        actual_rels = await _read_relationships(page)
        if item.get("codigo_60"):
            expected_code = str(item["codigo_60"]).strip()
            matching = [r for r in actual_rels if r["tipo"].upper() == RELATIONSHIP_TYPE.upper()]
            if not matching:
                result["diffs"].append({
                    "field": "Relacionamento",
                    "expected": f"{RELATIONSHIP_TYPE} / {expected_code}",
                    "actual": "(nenhum CÓDIGO ANTIGO)",
                })
            else:
                # Verificar código
                found_code = any(r["codigo"] == expected_code for r in matching)
                if not found_code:
                    result["diffs"].append({
                        "field": "Relacionamento (código)",
                        "expected": expected_code,
                        "actual": ", ".join(r["codigo"] for r in matching),
                    })
                # Verificar duplicatas
                if len(matching) > 1:
                    result["warnings"].append(
                        f"Relacionamento duplicado: {len(matching)} entradas de {RELATIONSHIP_TYPE}"
                    )

        # ── Mídias ──
        actual_media = await _read_media_count(page)
        doc_files = item.get("_doc_files", [])
        expected_docs = len(doc_files) if doc_files else 0
        if expected_docs > 0 and actual_media == 0:
            result["diffs"].append({
                "field": "Mídias",
                "expected": str(expected_docs),
                "actual": str(actual_media),
            })
        elif actual_media < expected_docs:
            result["warnings"].append(
                f"Mídias: esperado {expected_docs}, encontrado {actual_media}"
            )

        # ── PDM / Atributos ──
        pdm_data = await _read_pdm_and_attributes(page)
        if item.get("pdm") and not pdm_data.get("padronizado", True):
            result["diffs"].append({
                "field": "PDM",
                "expected": str(item["pdm"]),
                "actual": "(NÃO-PADRONIZADO)",
            })

        # Comparar atributos
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
            label = attr.get("label", f"Atrib_{i+1}")

            if exp_val.upper() != act_val.upper():
                result["diffs"].append({
                    "field": f"Atributo: {label}",
                    "expected": exp_val,
                    "actual": act_val or "(vazio)",
                })

    except Exception as e:
        result["status"] = "error"
        result["diffs"].append({"field": "ERRO", "expected": "", "actual": str(e)})
        log.error(f"  Erro ao verificar SIN {sin}: {e}")

    if result["diffs"]:
        result["status"] = "divergente"
    elif result["warnings"]:
        result["status"] = "ok_com_avisos"

    result["elapsed"] = round(time.time() - start, 1)
    return result


async def _voltar_worklist(page):
    """Volta para a worklist."""
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


async def run():
    log.info("=" * 60)
    log.info("Verificação Planilha vs Klassmatt — Início")
    log.info("=" * 60)

    # Determinar quais SINs verificar
    from_progress = "--from-progress" in sys.argv
    cli_sins = [s for s in sys.argv[1:] if not s.startswith("--")]

    # Ler Excel
    wb, items = load_excel()
    sin_data = {}
    for item in items:
        sin = str(item.get("sin", ""))
        if sin:
            sin_data[sin] = item

    if cli_sins:
        sins_to_verify = cli_sins
    elif from_progress:
        if PROGRESS_FILE.exists():
            progress = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
            sins_to_verify = [
                k for k, v in progress.get("items", {}).items()
                if v.get("status") == "ok"
            ]
        else:
            log.error("progress.json não encontrado")
            return
    else:
        sins_to_verify = list(sin_data.keys())

    log.info(f"SINs a verificar: {len(sins_to_verify)}")

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

    # Verificar cada SIN
    all_results = []
    counts = {"ok": 0, "ok_com_avisos": 0, "divergente": 0, "error": 0, "not_found": 0}

    for i, sin in enumerate(sins_to_verify):
        if sin not in sin_data:
            log.warning(f"SIN {sin} não encontrado na planilha — pulando")
            counts["not_found"] += 1
            continue

        log.info(f"[{i+1}/{len(sins_to_verify)}] Verificando SIN {sin}...")

        try:
            result = await verify_sin(page, sin, sin_data[sin])
            all_results.append(result)
            counts[result["status"]] = counts.get(result["status"], 0) + 1

            # Log resultado
            if result["status"] == "ok":
                log.info(f"  ✓ SIN {sin}: OK ({result['elapsed']:.1f}s)")
            elif result["status"] == "ok_com_avisos":
                log.info(f"  ~ SIN {sin}: OK com avisos ({result['elapsed']:.1f}s)")
                for w in result["warnings"]:
                    log.warning(f"    {w}")
            elif result["status"] == "divergente":
                log.warning(f"  ✗ SIN {sin}: DIVERGENTE ({result['elapsed']:.1f}s)")
                for d in result["diffs"]:
                    log.warning(f"    {d['field']}: esperado='{d['expected']}' | atual='{d['actual']}'")
                for w in result.get("warnings", []):
                    log.warning(f"    {w}")
            else:
                log.error(f"  ! SIN {sin}: ERRO ({result['elapsed']:.1f}s)")

            # Voltar para worklist
            await _voltar_worklist(page)
            await asyncio.sleep(3)

        except Exception as e:
            log.error(f"Erro fatal SIN {sin}: {e}")
            counts["error"] += 1
            all_results.append({
                "sin": sin, "status": "error",
                "diffs": [{"field": "ERRO FATAL", "expected": "", "actual": str(e)}],
                "warnings": [], "elapsed": 0,
            })
            try:
                await _voltar_worklist(page)
            except Exception:
                try:
                    for p in context.pages[1:]:
                        await p.close()
                    page = await context.new_page()
                    page.on("dialog", handle_dialog)
                    await page.goto(
                        "https://modec.klassmatt.com.br/MenuPrincipal.aspx",
                        wait_until="networkidle",
                    )
                    await _voltar_worklist(page)
                except Exception:
                    log.error("Recuperação falhou — encerrando")
                    break

    # Salvar relatório
    report = {
        "timestamp": datetime.now().isoformat(),
        "total": len(sins_to_verify),
        "counts": counts,
        "results": all_results,
    }
    REPORT_FILE.write_text(
        json.dumps(report, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    log.info(f"Relatório salvo em: {REPORT_FILE}")

    # Resumo
    log.info("=" * 60)
    log.info("RESUMO DA VERIFICAÇÃO")
    log.info("=" * 60)
    log.info(f"  Total verificados: {len(all_results)}")
    log.info(f"  OK:                {counts['ok']}")
    log.info(f"  OK com avisos:     {counts.get('ok_com_avisos', 0)}")
    log.info(f"  Divergentes:       {counts['divergente']}")
    log.info(f"  Erros:             {counts['error']}")
    log.info(f"  Não encontrados:   {counts.get('not_found', 0)}")

    if counts["divergente"] > 0:
        log.info("")
        log.info("SINs DIVERGENTES:")
        for r in all_results:
            if r["status"] == "divergente":
                diffs_str = ", ".join(d["field"] for d in r["diffs"])
                log.info(f"  SIN {r['sin']}: {diffs_str}")

    log.info("=" * 60)

    await context.close()
    await pw.stop()


if __name__ == "__main__":
    asyncio.run(run())
