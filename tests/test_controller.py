from __future__ import annotations

from pathlib import Path

from document_automation_studio.core.controller import ApplicationController
from document_automation_studio.models.config_models import ProcessingSettings


def test_application_controller_loads_default_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    controller = ApplicationController(config_path=config_path)

    assert controller.config.application_name == "Document Automation Studio"
    assert controller.config.processing.include_subfolders is True
    assert controller.config.processing.max_workers == 4


def test_application_controller_updates_processing_settings(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    controller = ApplicationController(config_path=config_path)
    new_settings = ProcessingSettings(
        input_folder=str(tmp_path / "input"),
        output_folder=str(tmp_path / "output"),
        include_subfolders=False,
        preserve_folder_structure=False,
        max_workers=2,
    )

    controller.update_processing_settings(new_settings)

    assert controller.config.processing.input_folder == str(tmp_path / "input")
    assert controller.config.processing.include_subfolders is False
    assert controller.processor.settings.max_workers == 2
    assert config_path.exists()
