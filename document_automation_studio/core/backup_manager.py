from __future__ import annotations

import logging
import shutil
from pathlib import Path

logger = logging.getLogger(__name__)


class BackupManager:
    """Manage backups for files before processing."""

    def __init__(self, backup_root: Path) -> None:
        self.backup_root = backup_root
        self.backup_root.mkdir(parents=True, exist_ok=True)

    def create_backup_folder(self, input_root: Path) -> None:
        backup_input_root = self.backup_root / input_root.name
        backup_input_root.mkdir(parents=True, exist_ok=True)
        logger.info("Backup folder ensured at %s", backup_input_root)

    def create_backup(self, source_path: Path, input_root: Path) -> Path:
        relative = source_path.relative_to(input_root)
        destination = self.backup_root / input_root.name / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination)
        logger.info("Backup created for %s at %s", source_path, destination)
        return destination

    def restore_backup(self, target_root: Path) -> None:
        if not self.backup_root.exists():
            logger.warning("No backups available to restore")
            return
        for backup_file in self.backup_root.rglob("*"):
            if backup_file.is_file():
                relative = backup_file.relative_to(self.backup_root)
                destination = target_root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(backup_file, destination)
                logger.info("Restored backup %s to %s", backup_file, destination)

    def delete_backup(self) -> None:
        if self.backup_root.exists():
            shutil.rmtree(self.backup_root)
            logger.info("Deleted all backups at %s", self.backup_root)
