"""Leitura do Excel e coloração de células para tracking."""

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import PatternFill

from config import EXCEL_PATH, EXCEL_COLUMNS, MAX_ATTRIBUTES, DOCUMENTS_DIR
from logger import log

# Cores para pintar linhas (compatível com o fluxo do André)
FILL_GREEN = PatternFill(start_color="00FF00", fill_type="solid")   # sucesso
FILL_RED = PatternFill(start_color="FF0000", fill_type="solid")     # erro
FILL_ORANGE = PatternFill(start_color="FFA500", fill_type="solid")  # duplicidade
FILL_YELLOW = PatternFill(start_color="FFFF00", fill_type="solid")  # needs_review


def load_excel(path: Path | None = None) -> tuple[openpyxl.Workbook, list[dict[str, Any]]]:
    """Lê a planilha e retorna (workbook, lista de dicts por linha).

    Cada dict tem as chaves definidas em EXCEL_COLUMNS + Atrib_N_Valor.
    Também inclui '_row' com o número da linha no Excel (para colorir).
    """
    path = path or EXCEL_PATH
    log.info(f"Abrindo Excel: {path}")
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    # Encontrar header row (primeira linha com dados nas colunas esperadas)
    headers = {}
    header_row = 1
    for row_idx in range(1, min(ws.max_row, 10) + 1):
        row_vals = [cell.value for cell in ws[row_idx]]
        if "SIN" in row_vals or any(
            v in row_vals for v in EXCEL_COLUMNS.values()
        ):
            headers = {
                ws.cell(row=row_idx, column=col).value: col
                for col in range(1, ws.max_column + 1)
                if ws.cell(row=row_idx, column=col).value
            }
            header_row = row_idx
            break

    if not headers:
        raise ValueError("Não foi possível encontrar cabeçalho no Excel")

    log.info(f"Cabeçalho encontrado na linha {header_row}: {list(headers.keys())[:10]}...")

    items = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        row_data = {"_row": row_idx}

        # Campos padrão
        for key, col_name in EXCEL_COLUMNS.items():
            col_idx = headers.get(col_name)
            row_data[key] = ws.cell(row=row_idx, column=col_idx).value if col_idx else None

        # Atributos técnicos (Atrib_1_Valor ... Atrib_30_Valor)
        attrs = []
        for n in range(1, MAX_ATTRIBUTES + 1):
            col_name = f"Atrib_{n}_Valor"
            col_idx = headers.get(col_name)
            val = ws.cell(row=row_idx, column=col_idx).value if col_idx else None
            attrs.append(val)
        row_data["attributes"] = attrs

        # Ignorar linhas completamente vazias
        if row_data.get("sin") is None:
            continue

        items.append(row_data)

    log.info(f"Total de itens lidos: {len(items)}")
    return wb, items


def color_row(wb: openpyxl.Workbook, row: int, status: str) -> None:
    """Pinta a célula A da linha com a cor correspondente ao status."""
    ws = wb.active
    fill = {
        "ok": FILL_GREEN,
        "error": FILL_RED,
        "duplicate": FILL_ORANGE,
        "skipped": FILL_ORANGE,
        "needs_review": FILL_YELLOW,
    }.get(status, FILL_RED)
    ws.cell(row=row, column=1).fill = fill


def save_excel(wb: openpyxl.Workbook, path: Path | None = None) -> None:
    """Salva o workbook (com as cores atualizadas)."""
    path = path or EXCEL_PATH
    wb.save(path)
    log.debug(f"Excel salvo: {path}")


def validate_documents(items: list[dict]) -> list[dict]:
    """Valida se os arquivos de documento existem no disco antes de processar.

    Retorna a lista de itens com campo '_missing_docs' preenchido.
    """
    for item in items:
        doc_str = item.get("documento")
        if not doc_str:
            item["_doc_files"] = []
            item["_missing_docs"] = []
            continue

        doc_names = [d.strip() for d in str(doc_str).split(";") if d.strip()]
        resolved = []
        missing = []
        for d in doc_names:
            full_path = DOCUMENTS_DIR / d
            if full_path.exists():
                resolved.append(str(full_path))
            else:
                # Tentar encontrar com extensão (ex: nome sem .pdf)
                matches = list(DOCUMENTS_DIR.glob(f"{d}.*"))
                if matches:
                    resolved.append(str(matches[0]))
                else:
                    missing.append(d)
        item["_doc_files"] = resolved
        item["_missing_docs"] = missing

        if missing:
            log.warning(
                f"SIN {item.get('sin')}: documentos não encontrados: {missing}"
            )

    return items
