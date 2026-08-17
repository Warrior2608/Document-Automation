from __future__ import annotations

import json
from pathlib import Path

from document_automation_studio.core.controller import ApplicationController


def test_controller_loads_rule_file_and_persists(tmp_path: Path) -> None:
    rule_file = tmp_path / "rules.json"
    rule_file.write_text(json.dumps({
        "text_replacements": [
            {"find": "{{Company}}", "replace": "ACME Corp", "case_sensitive": False}
        ]
    }), encoding="utf-8")

    controller = ApplicationController(config_path=tmp_path / "config.json")
    controller.load_rule_file(rule_file)

    assert controller.config.rule_files == [str(rule_file)]
    assert controller.rule_engine.rules.text_replacements[0].find == "{{Company}}"
    assert controller.config.active_profile == "Default"


def test_controller_rejects_invalid_rule_file(tmp_path: Path) -> None:
    rule_file = tmp_path / "bad_rules.json"
    rule_file.write_text(json.dumps([{"find": "{{Company}}", "replace": "ACME"}]), encoding="utf-8")

    controller = ApplicationController(config_path=tmp_path / "config.json")

    try:
        controller.load_rule_file(rule_file)
        assert False, "Expected ValueError for invalid rule file"
    except ValueError as error:
        assert "top level" in str(error)


def test_controller_profile_updates(tmp_path: Path) -> None:
    controller = ApplicationController(config_path=tmp_path / "config.json")
    controller.update_active_profile("Client A")

    assert controller.config.active_profile == "Client A"
    assert "Client A" in controller.config.profiles
