from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable

from document_automation_studio.config.manager import ConfigManager
from document_automation_studio.core.backup_manager import BackupManager
from document_automation_studio.core.batch_processor import BatchProcessor
from document_automation_studio.core.rename_engine import RenameEngine
from document_automation_studio.engine.preview_engine import PreviewEngine
from document_automation_studio.engine.rule_engine import RuleEngine
from document_automation_studio.models.config_models import AppConfig, ProcessingSettings
from document_automation_studio.models.preview_models import PreviewChange
from document_automation_studio.processors.branding_processor import BrandingProcessor, BrandingSettings
from document_automation_studio.processors.excel_processor import ExcelProcessor
from document_automation_studio.processors.metadata_processor import MetadataProcessor, MetadataUpdateResult, MetadataValues
from document_automation_studio.processors.word_processor import WordProcessor
from document_automation_studio.utils.logging_utils import configure_logging

logger = logging.getLogger(__name__)


class ApplicationController:
    """Coordinate GUI events, configuration, and the batch processing engine."""

    def __init__(self, config_path: str | Path | None = None) -> None:
        self.config_manager = ConfigManager(config_path=config_path)
        self.config = self.config_manager.load_config()
        self.logger = configure_logging(
            log_dir=self.config.logging.log_directory or Path("logs"),
            level=getattr(logging, self.config.logging.level.upper(), logging.INFO),
        )
        self.rule_engine = RuleEngine()
        self.preview_engine = PreviewEngine(self.rule_engine)
        self.rename_engine = RenameEngine()
        self.backup_manager = BackupManager(Path(self.config.processing.backup_directory) if self.config.processing.backup_directory else Path("backup"))
        self.word_processor = WordProcessor()
        self.excel_processor = ExcelProcessor()
        self.metadata_processor = MetadataProcessor()
        self.branding_processor = BrandingProcessor()
        self.metadata_values = self.metadata_processor.load_defaults()
        self.branding_settings = BrandingSettings()
        self.processor = BatchProcessor(settings=self.config.processing, worker_factory=self._worker_factory)

        if self.config.rule_files:
            try:
                self.load_rule_file(Path(self.config.rule_files[0]))
            except Exception:
                self.logger.exception("Failed to load configured rule file")

    def _worker_factory(self, path: Path) -> None:
        logger.info("Worker factory received file %s", path)
        output_root = Path(self.config.processing.output_folder or "output").resolve()
        output_root.mkdir(parents=True, exist_ok=True)
        input_root = Path(self.config.processing.input_folder).resolve() if self.config.processing.input_folder else None
        if path.suffix.lower() == ".docx":
            if not self.config.processing.process_word_files:
                logger.info("Skipping Word file because Word processing is disabled: %s", path)
                return
            saved_path = self.word_processor.process(
                source_path=path,
                output_root=output_root,
                preserve_folder_structure=self.config.processing.preserve_folder_structure,
                input_root=input_root,
                rule_set=self.rule_engine.rules if hasattr(self, "rule_engine") else None,
                logo_path=self.config.processing.logo_path,
                logo_width_in=self.config.processing.logo_width_inches,
                logo_height_in=self.config.processing.logo_height_inches,
                metadata_values=self.metadata_values,
                branding_settings=self.branding_settings,
            )
        elif path.suffix.lower() == ".xlsx":
            if not self.config.processing.process_excel_files:
                logger.info("Skipping Excel file because Excel processing is disabled: %s", path)
                return
            saved_path = self.excel_processor.process(
                source_path=path,
                output_root=output_root,
                preserve_folder_structure=self.config.processing.preserve_folder_structure,
                input_root=input_root,
                rule_set=self.rule_engine.rules,
            )
        else:
            logger.warning("No processor available for %s", path)
            return

        if self.config.processing.create_backups and input_root is not None:
            self.backup_manager.create_backup(path, input_root)

        renamed_path = self.rename_engine.apply_rename_rules(saved_path, self.rule_engine.rules.rename_patterns)
        if renamed_path != saved_path:
            saved_path.rename(renamed_path)
            logger.info("Renamed output file to %s", renamed_path)

    def save_config(self) -> None:
        self.config_manager.save_config(self.config)

    def update_processing_settings(self, settings: ProcessingSettings) -> None:
        self.config.processing = settings
        self.processor.settings = settings
        self.backup_manager = BackupManager(Path(self.config.processing.backup_directory) if self.config.processing.backup_directory else Path("backup"))
        self.save_config()

    def load_rule_file(self, rule_file_path: Path) -> None:
        self.rule_engine.load(rule_file_path)
        rule_file_text = str(rule_file_path)
        self.config.rule_files = [rule_file_text]
        self.save_config()

    def update_active_profile(self, profile_name: str) -> None:
        if profile_name not in self.config.profiles:
            self.config.profiles.append(profile_name)
        self.config.active_profile = profile_name
        self.save_config()

    def create_backup(self, input_root: Path) -> None:
        if not self.config.processing.create_backups:
            return
        self.backup_manager.create_backup_folder(input_root)

    def restore_backup(self, target_root: Path) -> None:
        self.backup_manager.restore_backup(target_root)

    def delete_backup(self) -> None:
        self.backup_manager.delete_backup()

    def scan_input_files(self) -> list[Path]:
        return self.processor.scan_files(Path(self.config.processing.input_folder))

    def run_batch(self, progress_callback: Callable[[int, int], None] | None = None) -> None:
        files = self.scan_input_files()
        self.processor.run(files, progress_callback=progress_callback)

    def set_metadata_values(self, values: MetadataValues) -> None:
        self.metadata_values = values

    def load_metadata_defaults(self) -> MetadataValues:
        self.metadata_values = self.metadata_processor.load_defaults()
        return self.metadata_values

    def save_metadata_defaults(self, values: MetadataValues) -> Path:
        return self.metadata_processor.save_defaults(values)

    def set_branding_settings(self, settings: BrandingSettings) -> None:
        self.branding_settings = settings

    def apply_metadata_to_documents(self, values: MetadataValues | None = None) -> list[MetadataUpdateResult]:
        selected_values = values or self.metadata_values
        files = self.scan_input_files()
        return self.metadata_processor.apply_to_documents(files, selected_values)

    def request_cancel(self) -> None:
        self.processor.cancel()

    def preview_changes(self, max_items: int = 20) -> list[PreviewChange]:
        files = self.scan_input_files()
        return self.preview_engine.preview_files(
            files,
            rule_set=self.rule_engine.rules,
            input_root=Path(self.config.processing.input_folder) if self.config.processing.input_folder else None,
            max_items=max_items,
            process_word_files=self.config.processing.process_word_files,
            process_excel_files=self.config.processing.process_excel_files,
            logo_path=self.config.processing.logo_path,
        )
