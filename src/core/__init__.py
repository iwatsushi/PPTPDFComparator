"""Core logic for document comparison (GUI-independent)."""

try:
    from src.core.document import Document, Page
    from src.core.exclusion_zone import ExclusionZone
    from src.core.page_matcher import PageMatcher, MatchResult
    from src.core.image_comparator import ImageComparator, DiffResult
    from src.core.session import Session
except ImportError:
    from core.document import Document, Page
    from core.exclusion_zone import ExclusionZone
    from core.page_matcher import PageMatcher, MatchResult
    from core.image_comparator import ImageComparator, DiffResult
    from core.session import Session

__all__ = [
    "Document",
    "Page",
    "ExclusionZone",
    "PageMatcher",
    "MatchResult",
    "ImageComparator",
    "DiffResult",
    "Session",
]
