from __future__ import annotations

import json
from pathlib import Path

from document_automation_studio.config.manager import ConfigManager
from document_automation_studio.models.config_models import AppConfig, ProcessingSettings


def test_default_config_contains_expected_sections(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    manager = ConfigManager(config_path=config_path)

    config = manager.create_default_config()

    assert isinstance(config, AppConfig)
    assert config.application_name == "Document Automation Studio"
    assert isinstance(config.processing, ProcessingSettings)
    assert config.processing.include_subfolders is True


def test_save_and_load_round_trip(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    manager = ConfigManager(config_path=config_path)
    config = manager.create_default_config()
    config.processing.input_folder = str(tmp_path / "input")
    config.processing.output_folder = str(tmp_path / "output")

    saved_path = manager.save_config(config)
    loaded_config = manager.load_config()

    assert saved_path == config_path
    assert loaded_config.processing.input_folder == str(tmp_path / "input")
    assert loaded_config.processing.output_folder == str(tmp_path / "output")
    assert config_path.exists()

    with config_path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    assert payload["application_name"] == "Document Automation Studio"
