"""Session management for saving and loading comparison state."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, List

try:
    from src.core.exclusion_zone import ExclusionZoneSet
    from src.core.page_matcher import MatchingResult
except ImportError:
    from core.exclusion_zone import ExclusionZoneSet
    from core.page_matcher import MatchingResult


@dataclass
class Session:
    """Represents a comparison session that can be saved and loaded."""

    left_document_path: Optional[str] = None
    right_document_path: Optional[str] = None
    matching_result: Optional[MatchingResult] = None
    exclusion_zones: ExclusionZoneSet = field(default_factory=ExclusionZoneSet)
    created_at: datetime = field(default_factory=datetime.now)
    modified_at: datetime = field(default_factory=datetime.now)
    notes: str = ""

    def save(self, file_path: str | Path) -> None:
        """Save session to a JSON file.

        Args:
            file_path: Path to save the session file
        """
        self.modified_at = datetime.now()

        data = {
            "version": "1.0",
            "left_document_path": self.left_document_path,
            "right_document_path": self.right_document_path,
            "matching_result": (
                self.matching_result.to_dict()
                if self.matching_result else None
            ),
            "exclusion_zones": self.exclusion_zones.to_dict(),
            "created_at": self.created_at.isoformat(),
            "modified_at": self.modified_at.isoformat(),
            "notes": self.notes,
        }

        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls, file_path: str | Path) -> Session:
        """Load session from a JSON file.

        Args:
            file_path: Path to the session file

        Returns:
            Loaded Session object
        """
        path = Path(file_path)

        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Parse matching result
        matching_result = None
        if data.get("matching_result"):
            matching_result = MatchingResult.from_dict(data["matching_result"])

        # Parse exclusion zones
        exclusion_zones = ExclusionZoneSet()
        if data.get("exclusion_zones"):
            exclusion_zones = ExclusionZoneSet.from_dict(data["exclusion_zones"])

        return cls(
            left_document_path=data.get("left_document_path"),
            right_document_path=data.get("right_document_path"),
            matching_result=matching_result,
            exclusion_zones=exclusion_zones,
            created_at=datetime.fromisoformat(data.get("created_at", datetime.now().isoformat())),
            modified_at=datetime.fromisoformat(data.get("modified_at", datetime.now().isoformat())),
            notes=data.get("notes", ""),
        )

    def has_documents(self) -> bool:
        """Check if both documents are set."""
        return bool(self.left_document_path and self.right_document_path)

    def clear(self) -> None:
        """Clear session state."""
        self.left_document_path = None
        self.right_document_path = None
        self.matching_result = None
        self.exclusion_zones = ExclusionZoneSet()
        self.notes = ""
        self.modified_at = datetime.now()
