"""Persistência de progresso em JSON para retomada após falha."""

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from config import PROGRESS_FILE
from logger import log


def load_progress() -> dict[str, Any]:
    """Carrega estado de progresso do arquivo JSON."""
    if PROGRESS_FILE.exists():
        data = json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
        log.info(f"Progresso carregado: {len(data.get('items', {}))} itens processados")
        return data
    return {"started_at": datetime.now().isoformat(), "items": {}}


def save_progress(progress: dict[str, Any]) -> None:
    """Salva estado de progresso no arquivo JSON."""
    progress["updated_at"] = datetime.now().isoformat()
    PROGRESS_FILE.write_text(
        json.dumps(progress, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def mark_item(progress: dict, sin: str, status: str, error: str = "") -> None:
    """Marca um item com status (ok, error, skipped, duplicate)."""
    progress["items"][sin] = {
        "status": status,
        "timestamp": datetime.now().isoformat(),
        "error": error,
    }
    save_progress(progress)
    log.info(f"SIN {sin} → {status}" + (f" ({error})" if error else ""))


def is_processed(progress: dict, sin: str) -> bool:
    """Verifica se um item já foi processado com sucesso."""
    item = progress.get("items", {}).get(sin)
    return item is not None and item["status"] == "ok"
