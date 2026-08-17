from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class PreviewChange:
    """Represents a single preview change entry."""

    file_path: Path
    change_type: str
    old_value: str
    new_value: str
    details: str = ""
