from __future__ import annotations

import json

from pathlib import Path

from document_automation_studio.engine.rule_engine import RuleEngine


def test_rule_engine_rejects_invalid_json_structure(tmp_path: Path) -> None:
    rule_file = tmp_path / "rules.json"
    rule_file.write_text(json.dumps([{"find": "A", "replace": "B"}]), encoding="utf-8")

    engine = RuleEngine()

    try:
        engine.load(rule_file)
        assert False, "Expected ValueError for invalid JSON structure"
    except ValueError as error:
        assert "top level" in str(error)


def test_rule_engine_rejects_csv_missing_fields(tmp_path: Path) -> None:
    rule_file = tmp_path / "rules.csv"
    rule_file.write_text("type,find\ntext,{{Company}}\n", encoding="utf-8")

    engine = RuleEngine()

    try:
        engine.load(rule_file)
        assert False, "Expected ValueError for invalid CSV rule row"
    except ValueError as error:
        assert "must include 'find' and 'replace'" in str(error)
