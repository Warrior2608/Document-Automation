from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from document_automation_studio.models.config_models import AppConfig, LoggingSettings, ProcessingSettings

logger = logging.getLogger(__name__)


class ConfigManager:
    """Manage application configuration stored as JSON."""

    def __init__(self, config_path: str | Path | None = None) -> None:
        self.config_path = Path(config_path or "config.json")
        self.config_path.parent.mkdir(parents=True, exist_ok=True)

    def create_default_config(self) -> AppConfig:
        """Create a default application configuration."""
        return AppConfig(
            processing=ProcessingSettings(),
            logging=LoggingSettings(),
        )

    def save_config(self, config: AppConfig) -> Path:
        """Persist configuration to disk as JSON."""
        payload: dict[str, Any] = {
            "application_name": config.application_name,
            "processing": {
                "input_folder": config.processing.input_folder,
                "output_folder": config.processing.output_folder,
                "include_subfolders": config.processing.include_subfolders,
                "process_word_files": config.processing.process_word_files,
                "process_excel_files": config.processing.process_excel_files,
                "preserve_folder_structure": config.processing.preserve_folder_structure,
                "skip_unsupported_files": config.processing.skip_unsupported_files,
                "max_workers": config.processing.max_workers,
                "create_backups": config.processing.create_backups,
                "backup_directory": config.processing.backup_directory,
                "logo_path": config.processing.logo_path,
                "logo_width_inches": config.processing.logo_width_inches,
                "logo_height_inches": config.processing.logo_height_inches,
            },
            "logging": {
                "level": config.logging.level,
                "log_directory": config.logging.log_directory,
                "log_to_file": config.logging.log_to_file,
                "log_to_console": config.logging.log_to_console,
            },
            "profiles": config.profiles,
            "active_profile": config.active_profile,
            "rule_files": config.rule_files,
        }

        self.config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        logger.info("Configuration saved to %s", self.config_path)
        return self.config_path

    def load_config(self) -> AppConfig:
        """Load configuration from disk, falling back to defaults."""
        if not self.config_path.exists():
            logger.info("Configuration file not found; using defaults")
            return self.create_default_config()

        with self.config_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)

        processing = ProcessingSettings(
            input_folder=payload.get("processing", {}).get("input_folder", ""),
            output_folder=payload.get("processing", {}).get("output_folder", ""),
            include_subfolders=payload.get("processing", {}).get("include_subfolders", True),
            process_word_files=payload.get("processing", {}).get("process_word_files", True),
            process_excel_files=payload.get("processing", {}).get("process_excel_files", True),
            preserve_folder_structure=payload.get("processing", {}).get("preserve_folder_structure", True),
            skip_unsupported_files=payload.get("processing", {}).get("skip_unsupported_files", True),
            max_workers=payload.get("processing", {}).get("max_workers", 4),
            create_backups=payload.get("processing", {}).get("create_backups", True),
            backup_directory=payload.get("processing", {}).get("backup_directory", "backup"),
            logo_path=payload.get("processing", {}).get("logo_path", ""),
            logo_width_inches=float(payload.get("processing", {}).get("logo_width_inches", 1.8) or 1.8),
            logo_height_inches=float(payload.get("processing", {}).get("logo_height_inches", 0.55) or 0.55),
        )
        logging_settings = LoggingSettings(
            level=payload.get("logging", {}).get("level", "INFO"),
            log_directory=payload.get("logging", {}).get("log_directory", ""),
            log_to_file=payload.get("logging", {}).get("log_to_file", True),
            log_to_console=payload.get("logging", {}).get("log_to_console", True),
        )
        return AppConfig(
            application_name=payload.get("application_name", "Document Automation Studio"),
            processing=processing,
            logging=logging_settings,
            profiles=payload.get("profiles", ["Default"]),
            rule_files=payload.get("rule_files", []),
            active_profile=payload.get("active_profile", "Default"),
        )
