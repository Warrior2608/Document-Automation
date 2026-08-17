from __future__ import annotations

from pathlib import Path

from document_automation_studio.core.backup_manager import BackupManager
from document_automation_studio.core.rename_engine import RenameEngine


def test_backup_manager_creates_backups(tmp_path: Path) -> None:
    input_root = tmp_path / "input"
    input_root.mkdir()
    source_file = input_root / "document.docx"
    source_file.write_text("test", encoding="utf-8")

    backup_root = tmp_path / "backup"
    manager = BackupManager(backup_root)
    backup = manager.create_backup(source_file, input_root)

    assert backup.exists()
    assert backup.read_text(encoding="utf-8") == "test"


def test_rename_engine_applies_rules(tmp_path: Path) -> None:
    target_path = tmp_path / "CompanyName_Report.docx"
    engine = RenameEngine()

    renamed = engine.apply_rename_rules(target_path, {"CompanyName": "ACME"})

    assert renamed.name == "ACME_Report.docx"
