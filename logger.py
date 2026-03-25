"""Logging estruturado para o RPA Klassmatt."""

import logging
import sys
from pathlib import Path

from config import LOG_FILE, SHARED_LOG_FILE


def setup_logger(name: str = "klassmatt_rpa") -> logging.Logger:
    """Configura logger com saída para arquivo, console e shared."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Arquivo local
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Arquivo shared (redundância na rede)
    try:
        SHARED_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        sh = logging.FileHandler(SHARED_LOG_FILE, encoding="utf-8")
        sh.setLevel(logging.DEBUG)
        sh.setFormatter(fmt)
        logger.addHandler(sh)
    except OSError as e:
        # Se a shared não estiver acessível, continua só com local
        print(f"[WARN] Shared log não acessível ({e}) — usando apenas log local")

    # Console
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


log = setup_logger()
