from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass(slots=True)
class ProcessingSettings:
    """Configuration for batch processing behavior."""

    input_folder: str = ""
    output_folder: str = ""
    include_subfolders: bool = True
    process_word_files: bool = True
    process_excel_files: bool = True
    preserve_folder_structure: bool = True
    skip_unsupported_files: bool = True
    max_workers: int = 4
    create_backups: bool = True
    backup_directory: str = "backup"
    logo_path: str = ""
    logo_width_inches: float = 1.8
    logo_height_inches: float = 0.55


@dataclass(slots=True)
class LoggingSettings:
    """Configuration for logging behavior."""

    level: str = "INFO"
    log_directory: str = ""
    log_to_file: bool = True
    log_to_console: bool = True


@dataclass(slots=True)
class AppConfig:
    """Top-level application configuration."""

    application_name: str = "Document Automation Studio"
    processing: ProcessingSettings = field(default_factory=ProcessingSettings)
    logging: LoggingSettings = field(default_factory=LoggingSettings)
    profiles: List[str] = field(default_factory=lambda: ["Default"])
    rule_files: List[str] = field(default_factory=list)
    active_profile: str = "Default"
