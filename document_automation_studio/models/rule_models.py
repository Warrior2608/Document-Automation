from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass(slots=True)
class TextReplacementRule:
    """Represents a single text replacement rule."""

    find: str
    replace: str
    case_sensitive: bool = False
    whole_word: bool = False
    regex: bool = False


@dataclass(slots=True)
class RuleSet:
    """Structured rule set loaded from configuration files."""

    text_replacements: List[TextReplacementRule] = field(default_factory=list)
    excel_replacements: List[TextReplacementRule] = field(default_factory=list)
    excel_insert_rows: List[dict[str, object]] = field(default_factory=list)
    excel_insert_columns: List[dict[str, object]] = field(default_factory=list)
    excel_hyperlinks: List[dict[str, str]] = field(default_factory=list)
    rename_patterns: dict[str, str] = field(default_factory=dict)
