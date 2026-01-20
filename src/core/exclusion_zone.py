"""Exclusion zone model for ignoring specific areas during comparison."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Literal, Optional
from enum import Enum
import json


class AppliesTo(Enum):
    """Which document(s) the exclusion zone applies to."""
    LEFT = "left"
    RIGHT = "right"
    BOTH = "both"


@dataclass
class ExclusionZone:
    """Defines a rectangular area to exclude from comparison.

    Coordinates are normalized (0.0 to 1.0) relative to page dimensions.
    """

    x: float  # Left edge (0.0 = left, 1.0 = right)
    y: float  # Top edge (0.0 = top, 1.0 = bottom)
    width: float  # Width as fraction of page width
    height: float  # Height as fraction of page height
    name: str = ""  # Human-readable name (e.g., "Page Number", "Header")
    applies_to: AppliesTo = AppliesTo.BOTH
    enabled: bool = True

    def __post_init__(self):
        """Validate coordinates."""
        if not (0.0 <= self.x <= 1.0):
            raise ValueError(f"x must be between 0.0 and 1.0, got {self.x}")
        if not (0.0 <= self.y <= 1.0):
            raise ValueError(f"y must be between 0.0 and 1.0, got {self.y}")
        if not (0.0 <= self.width <= 1.0):
            raise ValueError(f"width must be between 0.0 and 1.0, got {self.width}")
        if not (0.0 <= self.height <= 1.0):
            raise ValueError(f"height must be between 0.0 and 1.0, got {self.height}")

        if isinstance(self.applies_to, str):
            self.applies_to = AppliesTo(self.applies_to)

    def to_pixels(
        self,
        page_width: int,
        page_height: int
    ) -> tuple[int, int, int, int]:
        """Convert normalized coordinates to pixel coordinates.

        Returns:
            (x, y, width, height) in pixels
        """
        px_x = int(self.x * page_width)
        px_y = int(self.y * page_height)
        px_w = int(self.width * page_width)
        px_h = int(self.height * page_height)
        return (px_x, px_y, px_w, px_h)

    def to_rect(
        self,
        page_width: int,
        page_height: int
    ) -> tuple[int, int, int, int]:
        """Convert to rectangle coordinates (x1, y1, x2, y2).

        Returns:
            (left, top, right, bottom) in pixels
        """
        px_x, px_y, px_w, px_h = self.to_pixels(page_width, page_height)
        return (px_x, px_y, px_x + px_w, px_y + px_h)

    @classmethod
    def from_pixels(
        cls,
        x: int,
        y: int,
        width: int,
        height: int,
        page_width: int,
        page_height: int,
        **kwargs
    ) -> ExclusionZone:
        """Create an ExclusionZone from pixel coordinates.

        Args:
            x, y, width, height: Pixel coordinates
            page_width, page_height: Page dimensions for normalization
            **kwargs: Additional arguments (name, applies_to, enabled)
        """
        return cls(
            x=x / page_width,
            y=y / page_height,
            width=width / page_width,
            height=height / page_height,
            **kwargs
        )

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "name": self.name,
            "applies_to": self.applies_to.value,
            "enabled": self.enabled,
        }

    @classmethod
    def from_dict(cls, data: dict) -> ExclusionZone:
        """Create from dictionary."""
        return cls(
            x=data["x"],
            y=data["y"],
            width=data["width"],
            height=data["height"],
            name=data.get("name", ""),
            applies_to=AppliesTo(data.get("applies_to", "both")),
            enabled=data.get("enabled", True),
        )


@dataclass
class ExclusionZoneSet:
    """Collection of exclusion zones with common presets."""

    zones: List[ExclusionZone] = field(default_factory=list)

    def add(self, zone: ExclusionZone) -> None:
        """Add an exclusion zone."""
        self.zones.append(zone)

    def remove(self, zone: ExclusionZone) -> None:
        """Remove an exclusion zone."""
        self.zones.remove(zone)

    def clear(self) -> None:
        """Remove all zones."""
        self.zones.clear()

    def get_zones_for(self, side: Literal["left", "right"]) -> List[ExclusionZone]:
        """Get zones that apply to a specific side.

        Args:
            side: "left" or "right"

        Returns:
            List of applicable ExclusionZones
        """
        target = AppliesTo.LEFT if side == "left" else AppliesTo.RIGHT
        return [
            z for z in self.zones
            if z.enabled and (z.applies_to == target or z.applies_to == AppliesTo.BOTH)
        ]

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            "zones": [z.to_dict() for z in self.zones]
        }

    @classmethod
    def from_dict(cls, data: dict) -> ExclusionZoneSet:
        """Create from dictionary."""
        zones = [ExclusionZone.from_dict(z) for z in data.get("zones", [])]
        return cls(zones=zones)

    # Common presets
    @classmethod
    def preset_page_number_bottom(cls) -> ExclusionZone:
        """Preset for page number at bottom center."""
        return ExclusionZone(
            x=0.4,
            y=0.95,
            width=0.2,
            height=0.05,
            name="Page Number (Bottom)",
            applies_to=AppliesTo.BOTH,
        )

    @classmethod
    def preset_page_number_bottom_right(cls) -> ExclusionZone:
        """Preset for page number at bottom right."""
        return ExclusionZone(
            x=0.85,
            y=0.95,
            width=0.15,
            height=0.05,
            name="Page Number (Bottom Right)",
            applies_to=AppliesTo.BOTH,
        )

    @classmethod
    def preset_header(cls) -> ExclusionZone:
        """Preset for header area."""
        return ExclusionZone(
            x=0.0,
            y=0.0,
            width=1.0,
            height=0.08,
            name="Header",
            applies_to=AppliesTo.BOTH,
        )

    @classmethod
    def preset_footer(cls) -> ExclusionZone:
        """Preset for footer area."""
        return ExclusionZone(
            x=0.0,
            y=0.92,
            width=1.0,
            height=0.08,
            name="Footer",
            applies_to=AppliesTo.BOTH,
        )

    @classmethod
    def preset_slide_number_ppt(cls) -> ExclusionZone:
        """Preset for PowerPoint slide number (bottom right)."""
        return ExclusionZone(
            x=0.9,
            y=0.93,
            width=0.1,
            height=0.07,
            name="Slide Number",
            applies_to=AppliesTo.BOTH,
        )
