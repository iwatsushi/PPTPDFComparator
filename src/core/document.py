"""Document abstraction for PDF and PowerPoint files."""

from __future__ import annotations

import tempfile
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Optional, List, Tuple
from enum import Enum

import numpy as np
from PIL import Image
import imagehash

if TYPE_CHECKING:
    from numpy.typing import NDArray


# PowerPoint instance cache for faster subsequent loads
_powerpoint_cache = {
    'instance': None,
    'initialized': False
}


def _wait_for_files(folder: Path, min_count: int, timeout: float = 10.0, interval: float = 0.1) -> List[Path]:
    """Wait for PNG files to appear in folder with polling instead of fixed sleep.

    Args:
        folder: Directory to watch
        min_count: Minimum number of files expected
        timeout: Maximum wait time in seconds
        interval: Polling interval in seconds

    Returns:
        List of found PNG files
    """
    start = time.time()
    while time.time() - start < timeout:
        png_files = []
        for pattern in ["**/*.PNG", "**/*.png"]:
            png_files.extend(folder.glob(pattern))
        png_files = list(set(png_files))
        if len(png_files) >= min_count:
            return png_files
        time.sleep(interval)
    return png_files


def _wait_for_file(path: Path, timeout: float = 5.0, min_size: int = 1000) -> bool:
    """Wait for a file to exist and have minimum size.

    Args:
        path: File path to wait for
        timeout: Maximum wait time in seconds
        min_size: Minimum file size in bytes

    Returns:
        True if file exists and meets size requirement
    """
    start = time.time()
    while time.time() - start < timeout:
        if path.exists() and path.stat().st_size >= min_size:
            return True
        time.sleep(0.05)
    return False


def _compute_phash_for_image(args: Tuple[int, Image.Image, int]) -> Tuple[int, imagehash.ImageHash]:
    """Compute pHash for a single image (for parallel processing).

    Args:
        args: Tuple of (index, thumbnail, hash_size)

    Returns:
        Tuple of (index, computed hash)
    """
    idx, thumbnail, hash_size = args
    return idx, imagehash.phash(thumbnail, hash_size=hash_size)


def _compute_phashes_parallel(pages: List['Page'], hash_size: int = 16, max_workers: int = 4) -> None:
    """Compute pHashes for all pages in parallel.

    Args:
        pages: List of Page objects with thumbnails
        hash_size: Hash size for pHash
        max_workers: Maximum number of parallel workers
    """
    # Prepare work items
    work_items = [
        (i, page.thumbnail, hash_size)
        for i, page in enumerate(pages)
        if page.thumbnail is not None and page.phash is None
    ]

    if not work_items:
        return

    # Use ThreadPoolExecutor for parallel computation
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_compute_phash_for_image, item): item[0] for item in work_items}
        for future in as_completed(futures):
            try:
                idx, phash = future.result()
                pages[idx].phash = phash
            except Exception:
                pass  # Skip failed hashes


def get_powerpoint_instance():
    """Get or create cached PowerPoint COM instance.

    Returns:
        Tuple of (powerpoint_app, need_to_close) where need_to_close indicates
        if caller should close the instance
    """
    import sys
    if sys.platform != 'win32':
        raise RuntimeError("PowerPoint COM is only available on Windows")

    import comtypes
    import comtypes.client

    if not _powerpoint_cache['initialized']:
        comtypes.CoInitialize()
        _powerpoint_cache['initialized'] = True

    if _powerpoint_cache['instance'] is None:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1
        _powerpoint_cache['instance'] = powerpoint
        return powerpoint, False  # Don't close - it's cached

    return _powerpoint_cache['instance'], False


def close_powerpoint_cache():
    """Close cached PowerPoint instance."""
    if _powerpoint_cache['instance'] is not None:
        try:
            _powerpoint_cache['instance'].Quit()
        except Exception:
            pass
        _powerpoint_cache['instance'] = None


class DocumentType(Enum):
    """Supported document types."""
    PDF = "pdf"
    PPTX = "pptx"
    PPT = "ppt"


@dataclass
class Page:
    """Represents a single page/slide in a document."""

    index: int
    thumbnail: Optional[Image.Image] = None
    full_image: Optional[Image.Image] = None
    phash: Optional[imagehash.ImageHash] = None
    _cached_array: Optional[NDArray] = field(default=None, repr=False)

    @property
    def thumbnail_array(self) -> Optional[NDArray]:
        """Get thumbnail as numpy array."""
        if self.thumbnail is None:
            return None
        if self._cached_array is None:
            self._cached_array = np.array(self.thumbnail)
        return self._cached_array

    def compute_phash(self, hash_size: int = 16) -> imagehash.ImageHash:
        """Compute perceptual hash for this page."""
        if self.phash is None and self.thumbnail is not None:
            self.phash = imagehash.phash(self.thumbnail, hash_size=hash_size)
        return self.phash

    def get_full_image_array(self) -> Optional[NDArray]:
        """Get full resolution image as numpy array."""
        if self.full_image is None:
            return None
        return np.array(self.full_image)


@dataclass
class Document:
    """Represents a document (PDF or PowerPoint) as a collection of pages."""

    path: Path
    doc_type: DocumentType
    pages: List[Page] = field(default_factory=list)
    _loaded: bool = False

    @classmethod
    def from_file(cls, file_path: str | Path) -> Document:
        """Create a Document from a file path."""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

        suffix = path.suffix.lower()
        if suffix == ".pdf":
            doc_type = DocumentType.PDF
        elif suffix == ".pptx":
            doc_type = DocumentType.PPTX
        elif suffix == ".ppt":
            doc_type = DocumentType.PPT
        else:
            raise ValueError(f"Unsupported file type: {suffix}")

        return cls(path=path, doc_type=doc_type)

    @property
    def name(self) -> str:
        """Get document filename."""
        return self.path.name

    @property
    def page_count(self) -> int:
        """Get number of pages."""
        return len(self.pages)

    @property
    def is_loaded(self) -> bool:
        """Check if document has been loaded."""
        return self._loaded

    def load(
        self,
        thumbnail_size: tuple[int, int] = (800, 600),  # 高解像度サムネイル
        full_dpi: int = 200,  # 高DPI
        progress_callback: Optional[callable] = None
    ) -> None:
        """Load all pages from the document.

        Args:
            thumbnail_size: Size for thumbnail images (width, height)
            full_dpi: DPI for full resolution images
            progress_callback: Optional callback(current, total) for progress
        """
        if self._loaded:
            return

        if self.doc_type == DocumentType.PDF:
            self._load_pdf(thumbnail_size, full_dpi, progress_callback)
        elif self.doc_type in (DocumentType.PPTX, DocumentType.PPT):
            self._load_pptx(thumbnail_size, full_dpi, progress_callback)

        self._loaded = True

    def _load_pdf(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Load pages from PDF file using PyMuPDF.

        Optimizations:
        - Parallel pHash computation for many pages
        """
        import fitz  # PyMuPDF

        doc = fitz.open(str(self.path))
        total = len(doc)
        self.pages = []

        # Calculate zoom for thumbnail (72 DPI base, we want ~216 DPI equivalent)
        zoom = 3.0  # 3x zoom for better quality
        mat = fitz.Matrix(zoom, zoom)

        # First pass: load all thumbnails
        for i in range(total):
            if progress_callback:
                progress_callback(i + 1, total)

            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat)

            # Convert to PIL Image
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

            # Create thumbnail
            thumbnail = img.copy()
            thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

            page_obj = Page(index=i, thumbnail=thumbnail)
            self.pages.append(page_obj)

        doc.close()

        # Second pass: compute pHashes in parallel
        _compute_phashes_parallel(self.pages, hash_size=16, max_workers=4)
        # Note: Full images are loaded on-demand to save memory

    def _load_pptx(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Load pages from PowerPoint file."""
        from pptx import Presentation

        prs = Presentation(str(self.path))
        total = len(prs.slides)

        # Try different conversion methods in order of preference
        conversion_methods = [
            ("PowerPoint COM", self._convert_pptx_via_com),
            ("LibreOffice", self._convert_pptx_via_pdf),
        ]

        for method_name, method in conversion_methods:
            try:
                print(f"[DEBUG] Trying PPTX conversion via {method_name}...")
                method(thumbnail_size, full_dpi, progress_callback)
                print(f"[DEBUG] PPTX conversion via {method_name} succeeded")
                return
            except Exception as e:
                print(f"[DEBUG] {method_name} failed: {e}")
                continue

        # Fallback: create placeholder pages
        print("[DEBUG] All conversion methods failed, using placeholders")
        self.pages = []
        for i in range(total):
            if progress_callback:
                progress_callback(i + 1, total)

            # Create a placeholder thumbnail with text
            thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
            page = Page(index=i, thumbnail=thumbnail)
            page.compute_phash()
            self.pages.append(page)

    def _convert_pptx_via_com(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Convert PPTX to images using PowerPoint COM automation (Windows only).

        Optimizations:
        - Cached PowerPoint instance (faster subsequent loads)
        - Polling instead of fixed sleep (faster)
        - Parallel pHash computation (faster for many slides)
        - Optimized export resolution (1280px width for thumbnails)
        """
        import sys
        if sys.platform != 'win32':
            raise RuntimeError("PowerPoint COM is only available on Windows")

        try:
            import comtypes.client
        except ImportError:
            raise RuntimeError("comtypes not installed. Run: pip install comtypes")

        presentation = None
        powerpoint = None
        should_close_ppt = True

        try:
            # Use cached PowerPoint instance
            powerpoint, should_close_ppt = get_powerpoint_instance()

            # Open presentation (no fixed sleep - PowerPoint is already running)
            presentation = powerpoint.Presentations.Open(
                str(self.path.absolute()),
                ReadOnly=True,
                Untitled=False,
                WithWindow=True
            )

            total = presentation.Slides.Count
            self.pages = []

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                # Method 1: Bulk export via SaveAs (faster)
                export_folder = tmpdir_path / "slides"
                export_folder.mkdir(exist_ok=True)

                try:
                    export_path = str(export_folder / "slide.png")
                    presentation.SaveAs(export_path, 18)  # 18 = ppSaveAsPNG

                    # Use polling instead of fixed sleep
                    png_files = _wait_for_files(export_folder, total, timeout=15.0)

                    # Sort by slide number
                    def extract_number(path: Path) -> int:
                        match = re.search(r'(\d+)', path.stem)
                        return int(match.group(1)) if match else 0

                    png_files = sorted(set(png_files), key=extract_number)

                    if png_files and len(png_files) >= total:
                        # Load images without computing pHash yet
                        for i, png_file in enumerate(png_files[:total]):
                            if progress_callback:
                                progress_callback(i + 1, total)

                            img = Image.open(png_file)
                            thumbnail = img.copy()
                            thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

                            page = Page(index=i, thumbnail=thumbnail)
                            self.pages.append(page)

                        # Compute pHashes in parallel
                        _compute_phashes_parallel(self.pages, hash_size=16, max_workers=4)
                        return  # Success

                except Exception:
                    pass  # Fall through to individual export

                # Method 2: Individual slide export (fallback)
                # Use lower resolution for speed (1280px instead of 1920px * dpi factor)
                export_width = 1280

                for i in range(1, total + 1):
                    if progress_callback:
                        progress_callback(i, total)

                    slide = presentation.Slides(i)
                    img_path = tmpdir_path / f"slide_{i}.png"

                    try:
                        slide.Export(str(img_path), "PNG", export_width)

                        # Use polling instead of fixed sleep
                        if _wait_for_file(img_path, timeout=3.0, min_size=1000):
                            img = Image.open(img_path)
                            thumbnail = img.copy()
                            thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

                            page = Page(index=i - 1, thumbnail=thumbnail)
                            self.pages.append(page)
                        else:
                            # Create placeholder if export failed
                            thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                            page = Page(index=i - 1, thumbnail=thumbnail)
                            self.pages.append(page)
                    except Exception:
                        thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                        page = Page(index=i - 1, thumbnail=thumbnail)
                        self.pages.append(page)

                # Compute pHashes in parallel for all pages
                _compute_phashes_parallel(self.pages, hash_size=16, max_workers=4)

        finally:
            try:
                if presentation:
                    presentation.Close()
            except Exception:
                pass
            # Don't close PowerPoint if using cache
            if should_close_ppt and powerpoint:
                try:
                    powerpoint.Quit()
                except Exception:
                    pass

    def _convert_pptx_via_pdf(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Convert PPTX to PDF then to images using LibreOffice."""
        import subprocess
        import shutil

        # Check for LibreOffice
        soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice_path:
            raise RuntimeError("LibreOffice not found for PPTX conversion")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Convert PPTX to PDF
            subprocess.run([
                soffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(tmpdir),
                str(self.path)
            ], check=True, capture_output=True)

            pdf_path = tmpdir / self.path.with_suffix('.pdf').name

            if pdf_path.exists():
                # Load from the temporary PDF
                temp_doc = Document(path=pdf_path, doc_type=DocumentType.PDF)
                temp_doc._load_pdf(thumbnail_size, full_dpi, progress_callback)
                self.pages = temp_doc.pages
            else:
                raise RuntimeError("PDF conversion failed")

    def load_full_image(self, page_index: int, dpi: int = 150) -> Optional[Image.Image]:
        """Load full resolution image for a specific page on-demand.

        Args:
            page_index: Index of the page to load
            dpi: DPI for rendering

        Returns:
            Full resolution PIL Image
        """
        if page_index < 0 or page_index >= len(self.pages):
            raise IndexError(f"Page index out of range: {page_index}")

        page = self.pages[page_index]
        if page.full_image is not None:
            return page.full_image

        if self.doc_type == DocumentType.PDF:
            import fitz  # PyMuPDF
            doc = fitz.open(str(self.path))
            pdf_page = doc.load_page(page_index)

            # Calculate zoom based on DPI (72 is base DPI)
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = pdf_page.get_pixmap(matrix=mat)

            page.full_image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            doc.close()
        else:
            # For PPTX, we'd need similar on-demand conversion
            # For now, use thumbnail scaled up
            if page.thumbnail:
                page.full_image = page.thumbnail.copy()

        return page.full_image

    def compute_all_hashes(self, hash_size: int = 16) -> None:
        """Compute perceptual hashes for all pages."""
        for page in self.pages:
            page.compute_phash(hash_size)
