from __future__ import annotations

import logging
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable, Iterable

from document_automation_studio.models.config_models import ProcessingSettings

logger = logging.getLogger(__name__)


class BatchProcessor:
    """Batch processing engine for document automation workflows."""

    def __init__(self, settings: ProcessingSettings, worker_factory: Callable[[Path], None]) -> None:
        self.settings = settings
        self.worker_factory = worker_factory
        self._cancel_requested = False

    def scan_files(self, folder: Path) -> list[Path]:
        """Scan the input folder for supported Word and Excel files."""
        supported_extensions = set()
        if self.settings.process_word_files:
            supported_extensions.add("docx")
        if self.settings.process_excel_files:
            supported_extensions.add("xlsx")

        file_list: list[Path] = []
        if not folder.exists() or not folder.is_dir():
            logger.warning("Input folder does not exist: %s", folder)
            return file_list

        for path in folder.rglob("*") if self.settings.include_subfolders else folder.glob("*"):
            if path.is_file() and path.suffix.lower().lstrip(".") in supported_extensions:
                file_list.append(path)
            elif path.is_file() and not self.settings.skip_unsupported_files:
                logger.debug("Skipping unsupported file: %s", path)
        if not supported_extensions:
            logger.warning("No file types enabled for scanning")
        logger.info("Found %d supported files", len(file_list))
        return sorted(file_list)

    def run(self, files: Iterable[Path], progress_callback: Callable[[int, int], None] | None = None) -> None:
        """Execute batch processing with thread pooling and progress updates."""
        file_paths = list(files)
        if not file_paths:
            logger.info("No files to process")
            return

        total = len(file_paths)
        logger.info("Starting batch job for %d files", total)

        with ThreadPoolExecutor(max_workers=self.settings.max_workers) as executor:
            futures = []
            for index, file_path in enumerate(file_paths, start=1):
                if self._cancel_requested:
                    logger.warning("Batch cancellation requested before scheduling %s", file_path)
                    break
                futures.append(executor.submit(self._process_file, file_path, index, total, progress_callback))

            for future in futures:
                if self._cancel_requested:
                    logger.warning("Batch cancellation requested during processing")
                    break
                future.result()

        logger.info("Batch job completed")

    def cancel(self) -> None:
        """Request cancellation of the current batch operation."""
        self._cancel_requested = True

    def _process_file(self, path: Path, index: int, total: int, progress_callback: Callable[[int, int], None] | None) -> None:
        if self._cancel_requested:
            logger.debug("Skipping file because cancellation requested: %s", path)
            return

        try:
            self.worker_factory(path)
            logger.info("Processed file: %s", path)
        except Exception as error:  # pragma: no cover
            logger.exception("Error processing file %s: %s", path, error)

        if progress_callback:
            progress_callback(index, total)
