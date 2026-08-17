from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional


def configure_logging(log_dir: str | Path | None = None, level: int = logging.INFO) -> logging.Logger:
    """Configure a reusable application logger with file and console handlers."""
    log_path = Path(log_dir or "logs")
    log_path.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("document_automation_studio")
    logger.setLevel(level)
    logger.propagate = False
    # Reset handlers to ensure the caller-specified log directory and level are used.
    if logger.handlers:
        for h in list(logger.handlers):
            logger.removeHandler(h)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(str(log_path / "document_automation_studio.log"), encoding="utf-8")
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger
