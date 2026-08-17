from __future__ import annotations

import json
from pathlib import Path

from document_automation_studio.engine.rule_engine import RuleEngine
from document_automation_studio.models.rule_models import TextReplacementRule


def test_rule_engine_loads_json(tmp_path: Path) -> None:
    rule_file = tmp_path / "rules.json"
    rule_file.write_text(json.dumps({
        "text_replacements": [
            {"find": "{{Company}}", "replace": "ACME Corp", "case_sensitive": False, "whole_word": False, "regex": False}
        ],
        "rename_patterns": {"CompanyName": "ACME"}
    }), encoding="utf-8")

    engine = RuleEngine()
    engine.load(rule_file)

    assert len(engine.rules.text_replacements) == 1
    assert engine.rules.rename_patterns["CompanyName"] == "ACME"


def test_rule_engine_applies_text_replacements_case_insensitive() -> None:
    engine = RuleEngine()
    replacement = TextReplacementRule(find="Acme", replace="ACME", case_sensitive=False)

    result = engine.apply_text_replacements("Welcome acme corp", [replacement])

    assert result == "Welcome ACME corp"
