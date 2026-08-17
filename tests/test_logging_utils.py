from __future__ import annotations

import logging
from pathlib import Path

from document_automation_studio.utils.logging_utils import configure_logging


def test_configure_logging_creates_handlers(tmp_path: Path) -> None:
    logger = configure_logging(log_dir=tmp_path, level=logging.DEBUG)

    assert logger.name == "document_automation_studio"
    assert logger.level == logging.DEBUG
    assert len(logger.handlers) >= 2

    log_file = tmp_path / "document_automation_studio.log"
    logger.info("logging test")

    assert log_file.exists()
