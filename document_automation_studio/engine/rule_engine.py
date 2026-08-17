from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
from typing import Iterable

from document_automation_studio.models.rule_models import RuleSet, TextReplacementRule

logger = logging.getLogger(__name__)


class RuleEngine:
    """Load and apply automation rules from JSON or CSV rule files."""

    def __init__(self) -> None:
        self.rules = RuleSet()

    def load(self, rule_file: Path) -> None:
        logger.info("Loading rule file %s", rule_file)
        if not rule_file.exists():
            raise FileNotFoundError(f"Rule file not found: {rule_file}")

        if rule_file.suffix.lower() == ".json":
            self.rules = self._load_json(rule_file)
        elif rule_file.suffix.lower() == ".csv":
            self.rules = self._load_csv(rule_file)
        else:
            raise ValueError(f"Unsupported rule file type: {rule_file.suffix}")

    def _load_json(self, rule_file: Path) -> RuleSet:
        with rule_file.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)

        if not isinstance(payload, dict):
            raise ValueError("Rule file JSON must contain an object at the top level")

        self._validate_json_list(payload, "text_replacements")
        self._validate_json_list(payload, "excel_replacements")
        self._validate_json_list(payload, "excel_insert_rows", force_dict=False)
        self._validate_json_list(payload, "excel_insert_columns", force_dict=False)
        self._validate_json_list(payload, "excel_hyperlinks", force_dict=False)
        if "rename_patterns" in payload and not isinstance(payload["rename_patterns"], dict):
            raise ValueError("rename_patterns must be an object mapping old names to new names")

        replacements = [
            self._parse_text_replacement(item, "text_replacements")
            for item in payload.get("text_replacements", []) or []
        ]
        excel_replacements = [
            self._parse_text_replacement(item, "excel_replacements")
            for item in payload.get("excel_replacements", []) or []
        ]

        return RuleSet(
            text_replacements=replacements,
            excel_replacements=excel_replacements,
            excel_insert_rows=payload.get("excel_insert_rows", []) or [],
            excel_insert_columns=payload.get("excel_insert_columns", []) or [],
            excel_hyperlinks=payload.get("excel_hyperlinks", []) or [],
            rename_patterns=payload.get("rename_patterns", {}) or {},
        )

    def _load_csv(self, rule_file: Path) -> RuleSet:
        rules = RuleSet()
        with rule_file.open("r", encoding="utf-8", newline="") as handle:
            reader = csv.DictReader(handle)
            for row_number, row in enumerate(reader, start=1):
                if not row.get("find") or not row.get("replace"):
                    raise ValueError(f"CSV rule row {row_number} must include 'find' and 'replace' fields")
                rule_type = row.get("type", "text").strip().lower()
                if rule_type == "text":
                    rules.text_replacements.append(
                        TextReplacementRule(
                            find=row["find"],
                            replace=row["replace"],
                            case_sensitive=row.get("case_sensitive", "false").lower() == "true",
                            whole_word=row.get("whole_word", "false").lower() == "true",
                            regex=row.get("regex", "false").lower() == "true",
                        )
                    )
                elif rule_type == "rename":
                    rules.rename_patterns[row["find"]] = row["replace"]
                else:
                    raise ValueError(f"Unsupported CSV rule type '{rule_type}' at row {row_number}")

        return rules

    def _validate_json_list(self, payload: dict[str, object], key: str, force_dict: bool = True) -> None:
        if key not in payload:
            return
        value = payload[key]
        if not isinstance(value, list):
            raise ValueError(f"{key} must be a list")
        for entry in value:
            if not isinstance(entry, dict):
                raise ValueError(f"Each entry in {key} must be an object")
            if force_dict:
                if "find" not in entry or "replace" not in entry:
                    raise ValueError(f"Each entry in {key} must include 'find' and 'replace'")
            elif key in {"excel_insert_rows", "excel_insert_columns", "excel_hyperlinks"}:
                if key == "excel_hyperlinks" and ("cell" not in entry or "url" not in entry):
                    raise ValueError("Each entry in excel_hyperlinks must include 'cell' and 'url'")

    def _parse_text_replacement(self, item: dict[str, object], key: str) -> TextReplacementRule:
        if not isinstance(item, dict):
            raise ValueError(f"Each entry in {key} must be an object")
        if "find" not in item or "replace" not in item:
            raise ValueError(f"Each entry in {key} must include 'find' and 'replace'")
        return TextReplacementRule(
            find=item["find"],
            replace=item["replace"],
            case_sensitive=item.get("case_sensitive", False),
            whole_word=item.get("whole_word", False),
            regex=item.get("regex", False),
        )

    def apply_text_replacements(self, text: str, replacements: Iterable[TextReplacementRule]) -> str:
        result = text
        for replacement in replacements:
            if replacement.regex:
                import re

                flags = 0 if replacement.case_sensitive else re.IGNORECASE
                result = re.sub(replacement.find, replacement.replace, result, flags=flags)
            else:
                if replacement.case_sensitive:
                    result = result.replace(replacement.find, replacement.replace)
                else:
                    result = self._replace_case_insensitive(result, replacement.find, replacement.replace)
        return result

    def _replace_case_insensitive(self, source: str, find_text: str, replace_text: str) -> str:
        import re

        pattern = re.compile(re.escape(find_text), re.IGNORECASE)
        return pattern.sub(replace_text, source)
