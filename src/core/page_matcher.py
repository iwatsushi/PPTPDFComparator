"""Page matching algorithm using perceptual hashing and optimal assignment."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Set
from enum import Enum

import numpy as np
from scipy.optimize import linear_sum_assignment
import imagehash

try:
    from src.core.document import Document, Page
except ImportError:
    from core.document import Document, Page


class MatchStatus(Enum):
    """Status of a page match."""
    MATCHED = "matched"  # Pages are matched
    UNMATCHED_LEFT = "unmatched_left"  # Left page has no match
    UNMATCHED_RIGHT = "unmatched_right"  # Right page has no match


@dataclass
class MatchResult:
    """Result of matching two pages."""

    left_index: Optional[int]  # None if unmatched right
    right_index: Optional[int]  # None if unmatched left
    status: MatchStatus
    similarity: float = 0.0  # 0.0 to 1.0, higher is more similar
    phash_distance: int = 0  # Hamming distance of perceptual hashes
    is_manual: bool = False  # True if manually set by user

    @property
    def has_difference(self) -> bool:
        """Check if matched pages have differences."""
        if self.status != MatchStatus.MATCHED:
            return True
        return self.similarity < 0.99


@dataclass
class MatchingResult:
    """Complete result of matching two documents."""

    matches: List[MatchResult] = field(default_factory=list)
    left_unmatched: Set[int] = field(default_factory=set)
    right_unmatched: Set[int] = field(default_factory=set)

    def get_match_for_left(self, left_index: int) -> Optional[MatchResult]:
        """Get match result for a left page index."""
        for m in self.matches:
            if m.left_index == left_index:
                return m
        return None

    def get_match_for_right(self, right_index: int) -> Optional[MatchResult]:
        """Get match result for a right page index."""
        for m in self.matches:
            if m.right_index == right_index:
                return m
        return None

    def get_matched_pairs(self) -> List[Tuple[int, int, float]]:
        """Get list of matched pairs as (left_index, right_index, similarity)."""
        pairs = []
        for m in self.matches:
            if m.status == MatchStatus.MATCHED and m.left_index is not None and m.right_index is not None:
                pairs.append((m.left_index, m.right_index, m.similarity))
        return pairs

    def set_manual_match(
        self,
        left_index: Optional[int],
        right_index: Optional[int]
    ) -> None:
        """Manually set or update a match.

        Args:
            left_index: Left page index (None to mark right as unmatched)
            right_index: Right page index (None to mark left as unmatched)
        """
        # Remove existing matches for these indices
        self.matches = [
            m for m in self.matches
            if m.left_index != left_index and m.right_index != right_index
        ]

        if left_index is not None:
            self.left_unmatched.discard(left_index)
        if right_index is not None:
            self.right_unmatched.discard(right_index)

        if left_index is not None and right_index is not None:
            # Create matched pair
            self.matches.append(MatchResult(
                left_index=left_index,
                right_index=right_index,
                status=MatchStatus.MATCHED,
                similarity=1.0,  # Will be recalculated
                is_manual=True,
            ))
        elif left_index is not None:
            # Mark left as unmatched
            self.left_unmatched.add(left_index)
            self.matches.append(MatchResult(
                left_index=left_index,
                right_index=None,
                status=MatchStatus.UNMATCHED_LEFT,
                is_manual=True,
            ))
        elif right_index is not None:
            # Mark right as unmatched
            self.right_unmatched.add(right_index)
            self.matches.append(MatchResult(
                left_index=None,
                right_index=right_index,
                status=MatchStatus.UNMATCHED_RIGHT,
                is_manual=True,
            ))

    def remove_manual_match(self, left_index: int, right_index: int) -> None:
        """Remove a manual match."""
        self.matches = [
            m for m in self.matches
            if not (m.left_index == left_index and m.right_index == right_index and m.is_manual)
        ]

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            "matches": [
                {
                    "left_index": m.left_index,
                    "right_index": m.right_index,
                    "status": m.status.value,
                    "similarity": m.similarity,
                    "phash_distance": m.phash_distance,
                    "is_manual": m.is_manual,
                }
                for m in self.matches
            ],
            "left_unmatched": list(self.left_unmatched),
            "right_unmatched": list(self.right_unmatched),
        }

    @classmethod
    def from_dict(cls, data: dict) -> MatchingResult:
        """Create from dictionary."""
        matches = [
            MatchResult(
                left_index=m["left_index"],
                right_index=m["right_index"],
                status=MatchStatus(m["status"]),
                similarity=m.get("similarity", 0.0),
                phash_distance=m.get("phash_distance", 0),
                is_manual=m.get("is_manual", False),
            )
            for m in data.get("matches", [])
        ]
        return cls(
            matches=matches,
            left_unmatched=set(data.get("left_unmatched", [])),
            right_unmatched=set(data.get("right_unmatched", [])),
        )


class PageMatcher:
    """Matches pages between two documents using perceptual hashing."""

    def __init__(
        self,
        phash_threshold: int = 20,
        position_weight: float = 0.1,
        hash_size: int = 16,
    ):
        """Initialize the matcher.

        Args:
            phash_threshold: Maximum hamming distance for candidate pairs
            position_weight: Weight for position penalty (0.0 to 1.0)
            hash_size: Size of perceptual hash
        """
        self.phash_threshold = phash_threshold
        self.position_weight = position_weight
        self.hash_size = hash_size

    def match(
        self,
        left_doc: Document,
        right_doc: Document,
        progress_callback: Optional[callable] = None
    ) -> MatchingResult:
        """Match pages between two documents.

        Uses a two-phase approach:
        1. Coarse matching with perceptual hashing to find candidates
        2. Hungarian algorithm for optimal assignment

        Args:
            left_doc: Left document
            right_doc: Right document
            progress_callback: Optional callback(current, total, message)

        Returns:
            MatchingResult with all matches
        """
        left_pages = left_doc.pages
        right_pages = right_doc.pages
        n_left = len(left_pages)
        n_right = len(right_pages)

        if n_left == 0 or n_right == 0:
            return self._handle_empty_documents(n_left, n_right)

        # Ensure all pages have hashes
        for page in left_pages + right_pages:
            page.compute_phash(self.hash_size)

        if progress_callback:
            progress_callback(0, n_left * n_right, "Computing similarity matrix...")

        # Phase 1: Build candidate matrix using pHash
        candidates = self._build_candidate_matrix(
            left_pages, right_pages, progress_callback
        )

        if progress_callback:
            progress_callback(n_left * n_right, n_left * n_right, "Running assignment...")

        # Phase 2: Hungarian algorithm
        result = self._solve_assignment(
            candidates, n_left, n_right, left_pages, right_pages
        )

        return result

    def _handle_empty_documents(
        self,
        n_left: int,
        n_right: int
    ) -> MatchingResult:
        """Handle case where one or both documents are empty."""
        result = MatchingResult()

        for i in range(n_left):
            result.left_unmatched.add(i)
            result.matches.append(MatchResult(
                left_index=i,
                right_index=None,
                status=MatchStatus.UNMATCHED_LEFT,
            ))

        for j in range(n_right):
            result.right_unmatched.add(j)
            result.matches.append(MatchResult(
                left_index=None,
                right_index=j,
                status=MatchStatus.UNMATCHED_RIGHT,
            ))

        return result

    def _build_candidate_matrix(
        self,
        left_pages: List[Page],
        right_pages: List[Page],
        progress_callback: Optional[callable]
    ) -> Dict[Tuple[int, int], Tuple[int, float]]:
        """Build matrix of candidate pairs using pHash.

        Returns:
            Dictionary mapping (left_idx, right_idx) to (hash_distance, similarity)
        """
        candidates = {}
        total = len(left_pages) * len(right_pages)
        count = 0

        for i, lp in enumerate(left_pages):
            for j, rp in enumerate(right_pages):
                count += 1
                if progress_callback and count % 100 == 0:
                    progress_callback(count, total, "Computing similarities...")

                if lp.phash is None or rp.phash is None:
                    continue

                # Compute hamming distance
                distance = lp.phash - rp.phash

                # Only consider pairs below threshold
                if distance <= self.phash_threshold:
                    # Convert distance to similarity (0 distance = 1.0 similarity)
                    max_distance = self.hash_size * self.hash_size
                    similarity = 1.0 - (distance / max_distance)
                    candidates[(i, j)] = (distance, similarity)

        return candidates

    def _solve_assignment(
        self,
        candidates: Dict[Tuple[int, int], Tuple[int, float]],
        n_left: int,
        n_right: int,
        left_pages: List[Page],
        right_pages: List[Page]
    ) -> MatchingResult:
        """Solve optimal assignment using Hungarian algorithm.

        Uses position penalty to prefer matches that maintain relative order.
        """
        # Build cost matrix
        # Use large cost for non-candidates
        INF_COST = 1e9
        n = max(n_left, n_right)
        cost_matrix = np.full((n, n), INF_COST)

        for (i, j), (distance, similarity) in candidates.items():
            # Base cost from hash distance
            base_cost = distance

            # Position penalty: prefer matches that maintain order
            # Normalized position difference
            pos_diff = abs(i / n_left - j / n_right) if n_left > 0 and n_right > 0 else 0
            position_penalty = pos_diff * self.position_weight * self.hash_size * self.hash_size

            cost_matrix[i, j] = base_cost + position_penalty

        # Solve assignment
        row_ind, col_ind = linear_sum_assignment(cost_matrix)

        # Build result
        result = MatchingResult()
        matched_left = set()
        matched_right = set()

        for i, j in zip(row_ind, col_ind):
            if i < n_left and j < n_right and cost_matrix[i, j] < INF_COST:
                distance, similarity = candidates.get((i, j), (0, 0.0))
                result.matches.append(MatchResult(
                    left_index=i,
                    right_index=j,
                    status=MatchStatus.MATCHED,
                    similarity=similarity,
                    phash_distance=distance,
                ))
                matched_left.add(i)
                matched_right.add(j)

        # Add unmatched pages
        for i in range(n_left):
            if i not in matched_left:
                result.left_unmatched.add(i)
                result.matches.append(MatchResult(
                    left_index=i,
                    right_index=None,
                    status=MatchStatus.UNMATCHED_LEFT,
                ))

        for j in range(n_right):
            if j not in matched_right:
                result.right_unmatched.add(j)
                result.matches.append(MatchResult(
                    left_index=None,
                    right_index=j,
                    status=MatchStatus.UNMATCHED_RIGHT,
                ))

        # Sort matches by left index (with unmatched at end)
        result.matches.sort(
            key=lambda m: (
                m.left_index if m.left_index is not None else n_left + (m.right_index or 0),
                m.right_index if m.right_index is not None else n_right
            )
        )

        return result

    def refine_with_ssim(
        self,
        result: MatchingResult,
        left_doc: Document,
        right_doc: Document,
        progress_callback: Optional[callable] = None
    ) -> None:
        """Refine similarity scores using SSIM for matched pairs.

        This is more accurate but slower than pHash.
        Modifies the result in-place.
        """
        from skimage.metrics import structural_similarity as ssim

        matched = [m for m in result.matches if m.status == MatchStatus.MATCHED]
        total = len(matched)

        for idx, match in enumerate(matched):
            if progress_callback:
                progress_callback(idx + 1, total, "Computing SSIM...")

            if match.left_index is None or match.right_index is None:
                continue

            left_img = left_doc.pages[match.left_index].thumbnail
            right_img = right_doc.pages[match.right_index].thumbnail

            if left_img is None or right_img is None:
                continue

            # Convert to same size grayscale
            left_gray = left_img.convert('L')
            right_gray = right_img.convert('L')

            # Resize to same dimensions
            target_size = (min(left_gray.width, right_gray.width),
                           min(left_gray.height, right_gray.height))
            left_gray = left_gray.resize(target_size)
            right_gray = right_gray.resize(target_size)

            # Compute SSIM
            left_arr = np.array(left_gray)
            right_arr = np.array(right_gray)
            score = ssim(left_arr, right_arr)

            match.similarity = score
