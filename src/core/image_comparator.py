"""Image comparison and difference highlighting."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, List, Optional, Tuple
import cv2
import numpy as np
from PIL import Image

if TYPE_CHECKING:
    from numpy.typing import NDArray

try:
    from src.core.exclusion_zone import ExclusionZone, ExclusionZoneSet
except ImportError:
    from core.exclusion_zone import ExclusionZone, ExclusionZoneSet


@dataclass
class DiffRegion:
    """A region of difference between two images."""

    x: int
    y: int
    width: int
    height: int
    area: int = 0
    intensity: float = 0.0  # Average difference intensity

    @property
    def rect(self) -> Tuple[int, int, int, int]:
        """Get as (x, y, w, h)."""
        return (self.x, self.y, self.width, self.height)

    @property
    def bounds(self) -> Tuple[int, int, int, int]:
        """Get as (x1, y1, x2, y2)."""
        return (self.x, self.y, self.x + self.width, self.y + self.height)


@dataclass
class DiffResult:
    """Result of comparing two images."""

    diff_score: float  # 0.0 = identical, 1.0 = completely different
    regions: List[DiffRegion] = field(default_factory=list)
    diff_image: Optional[Image.Image] = None  # Grayscale diff
    highlight_image: Optional[Image.Image] = None  # Original with highlights

    @property
    def has_differences(self) -> bool:
        """Check if there are any meaningful differences."""
        # Either diff_score > 1% OR at least one region detected
        return self.diff_score > 0.01 or len(self.regions) > 0

    @property
    def diff_count(self) -> int:
        """Number of difference regions."""
        return len(self.regions)


class ImageComparator:
    """Compares two images and highlights differences."""

    def __init__(
        self,
        threshold: int = 30,  # レンダリング差を吸収する閾値
        min_region_area: int = 100,  # ノイズを除去（小さすぎる差分は無視）
        highlight_color: Tuple[int, int, int] = (255, 0, 0),
        highlight_alpha: float = 0.5,
    ):
        """Initialize the comparator.

        Args:
            threshold: Pixel difference threshold (0-255)
            min_region_area: Minimum area for a difference region
            highlight_color: RGB color for highlighting (default: red)
            highlight_alpha: Opacity of highlight overlay (0.0-1.0)
        """
        self.threshold = threshold
        self.min_region_area = min_region_area
        self.highlight_color = highlight_color
        self.highlight_alpha = highlight_alpha

    def compare(
        self,
        img1: Image.Image,
        img2: Image.Image,
        exclusion_zones: Optional[List[ExclusionZone]] = None
    ) -> DiffResult:
        """Compare two images and find differences.

        Args:
            img1: First image (PIL Image)
            img2: Second image (PIL Image)
            exclusion_zones: Areas to ignore during comparison

        Returns:
            DiffResult with difference information and highlighted image
        """
        # Ensure same size
        target_size = (
            max(img1.width, img2.width),
            max(img1.height, img2.height)
        )

        if img1.size != target_size:
            img1 = img1.resize(target_size, Image.Resampling.LANCZOS)
        if img2.size != target_size:
            img2 = img2.resize(target_size, Image.Resampling.LANCZOS)

        # Convert to numpy arrays
        arr1 = np.array(img1.convert('RGB'))
        arr2 = np.array(img2.convert('RGB'))

        # Convert to grayscale for comparison
        gray1 = cv2.cvtColor(arr1, cv2.COLOR_RGB2GRAY)
        gray2 = cv2.cvtColor(arr2, cv2.COLOR_RGB2GRAY)

        # Compute absolute difference
        diff = cv2.absdiff(gray1, gray2)

        # Apply exclusion mask
        if exclusion_zones:
            mask = self._create_exclusion_mask(
                target_size[0], target_size[1], exclusion_zones
            )
            diff = cv2.bitwise_and(diff, diff, mask=cv2.bitwise_not(mask))

        # Threshold to binary
        _, thresh = cv2.threshold(diff, self.threshold, 255, cv2.THRESH_BINARY)

        # Find contours (difference regions)
        contours, _ = cv2.findContours(
            thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        # Filter and process regions
        regions = []
        for contour in contours:
            area = cv2.contourArea(contour)
            if area >= self.min_region_area:
                x, y, w, h = cv2.boundingRect(contour)

                # Calculate intensity in this region
                roi = diff[y:y+h, x:x+w]
                intensity = float(np.mean(roi)) / 255.0

                regions.append(DiffRegion(
                    x=x, y=y, width=w, height=h,
                    area=int(area), intensity=intensity
                ))

        # Calculate overall diff score
        total_pixels = target_size[0] * target_size[1]
        diff_pixels = np.count_nonzero(thresh)
        diff_score = diff_pixels / total_pixels if total_pixels > 0 else 0.0

        # Create diff image
        diff_pil = Image.fromarray(diff)

        # Create highlight overlay
        highlight = self._create_highlight_image(arr1, regions)

        return DiffResult(
            diff_score=diff_score,
            regions=regions,
            diff_image=diff_pil,
            highlight_image=highlight,
        )

    def _create_exclusion_mask(
        self,
        width: int,
        height: int,
        zones: List[ExclusionZone]
    ) -> NDArray:
        """Create a binary mask for exclusion zones.

        Returns:
            uint8 array where 255 = excluded, 0 = included
        """
        mask = np.zeros((height, width), dtype=np.uint8)

        for zone in zones:
            if not zone.enabled:
                continue

            x1, y1, x2, y2 = zone.to_rect(width, height)
            # Clamp to image bounds
            x1 = max(0, min(x1, width))
            y1 = max(0, min(y1, height))
            x2 = max(0, min(x2, width))
            y2 = max(0, min(y2, height))

            mask[y1:y2, x1:x2] = 255

        return mask

    def _create_highlight_image(
        self,
        base_image: NDArray,
        regions: List[DiffRegion]
    ) -> Image.Image:
        """Create image with difference regions highlighted.

        Args:
            base_image: RGB numpy array
            regions: List of difference regions

        Returns:
            PIL Image with highlights
        """
        result = base_image.copy()

        # Create overlay for each region
        for region in regions:
            x1, y1, x2, y2 = region.bounds

            # Draw semi-transparent highlight
            overlay = result[y1:y2, x1:x2].copy()
            highlight = np.full_like(overlay, self.highlight_color)

            blended = cv2.addWeighted(
                overlay, 1 - self.highlight_alpha,
                highlight, self.highlight_alpha,
                0
            )
            result[y1:y2, x1:x2] = blended

            # Draw border
            cv2.rectangle(
                result,
                (x1, y1), (x2, y2),
                self.highlight_color,
                2
            )

        return Image.fromarray(result)

    def create_side_by_side(
        self,
        img1: Image.Image,
        img2: Image.Image,
        diff_result: Optional[DiffResult] = None,
        gap: int = 10,
    ) -> Image.Image:
        """Create side-by-side comparison image.

        Args:
            img1: Left image
            img2: Right image (or highlighted version)
            diff_result: Optional diff result for annotations
            gap: Gap between images in pixels

        Returns:
            Combined image
        """
        # Use highlighted image if available
        if diff_result and diff_result.highlight_image:
            img2 = diff_result.highlight_image

        # Ensure same height
        max_height = max(img1.height, img2.height)
        if img1.height != max_height:
            ratio = max_height / img1.height
            img1 = img1.resize(
                (int(img1.width * ratio), max_height),
                Image.Resampling.LANCZOS
            )
        if img2.height != max_height:
            ratio = max_height / img2.height
            img2 = img2.resize(
                (int(img2.width * ratio), max_height),
                Image.Resampling.LANCZOS
            )

        # Create combined image
        total_width = img1.width + gap + img2.width
        combined = Image.new('RGB', (total_width, max_height), (255, 255, 255))
        combined.paste(img1, (0, 0))
        combined.paste(img2, (img1.width + gap, 0))

        return combined
