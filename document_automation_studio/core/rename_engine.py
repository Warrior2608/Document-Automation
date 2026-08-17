from __future__ import annotations

from pathlib import Path


class RenameEngine:
    """Apply rename patterns to output files."""

    def apply_rename_rules(self, target_path: Path, rename_patterns: dict[str, str]) -> Path:
        new_name = target_path.name
        for find_text, replace_text in rename_patterns.items():
            new_name = new_name.replace(find_text, replace_text)

        if new_name == target_path.name:
            return target_path

        return target_path.with_name(new_name)
