from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from document_automation_studio.engine.preview_engine import PreviewEngine
from document_automation_studio.engine.rule_engine import RuleEngine
from document_automation_studio.models.rule_models import RuleSet, TextReplacementRule


def test_preview_engine_insert_rows_and_columns(tmp_path: Path) -> None:
    source = tmp_path / "workbook.xlsx"
    workbook = Workbook()
    workbook.save(source)

    engine = PreviewEngine(RuleEngine())
    rule_set = RuleSet(
        excel_insert_rows=[{"sheet": "Sheet", "index": 2, "values": ["A", "B"]}],
        excel_insert_columns=[{"sheet": "Sheet", "index": 1, "values": ["X", "Y"]}],
    )

    preview_items = engine.preview_files([source], rule_set=rule_set, input_root=tmp_path, max_items=10)

    assert any(item.change_type == "Excel Insert Row" for item in preview_items)
    assert any(item.change_type == "Excel Insert Column" for item in preview_items)


def test_preview_engine_hyperlink_rules(tmp_path: Path) -> None:
    source = tmp_path / "workbook.xlsx"
    workbook = Workbook()
    workbook.save(source)

    engine = PreviewEngine(RuleEngine())
    rule_set = RuleSet(
        excel_hyperlinks=[{"sheet": "Sheet", "cell": "C3", "url": "https://example.com", "display": "Example"}]
    )

    preview_items = engine.preview_files([source], rule_set=rule_set, input_root=tmp_path, max_items=10)

    assert len(preview_items) == 1
    assert preview_items[0].change_type == "Excel Hyperlink"
    assert "C3" in preview_items[0].new_value
    assert "Example" in preview_items[0].new_value
