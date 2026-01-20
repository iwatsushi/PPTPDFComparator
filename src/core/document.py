"""Document abstraction for PDF and PowerPoint files."""

from __future__ import annotations

import io
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Optional, List
from enum import Enum
import hashlib

import numpy as np
from PIL import Image
import imagehash

if TYPE_CHECKING:
    from numpy.typing import NDArray


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
        """Load pages from PDF file using PyMuPDF."""
        import fitz  # PyMuPDF

        doc = fitz.open(str(self.path))
        total = len(doc)
        self.pages = []

        # Calculate zoom for thumbnail (72 DPI base, we want ~216 DPI equivalent)
        zoom = 3.0  # 3x zoom for better quality
        mat = fitz.Matrix(zoom, zoom)

        for i in range(total):
            if progress_callback:
                progress_callback(i + 1, total)

            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat)

            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Create thumbnail
            thumbnail = img.copy()
            thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

            page_obj = Page(index=i, thumbnail=thumbnail)
            page_obj.compute_phash()
            self.pages.append(page_obj)

        doc.close()
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
        """Convert PPTX to images using PowerPoint COM automation (Windows only)."""
        import sys
        import time
        if sys.platform != 'win32':
            raise RuntimeError("PowerPoint COM is only available on Windows")

        try:
            import comtypes.client
        except ImportError:
            raise RuntimeError("comtypes not installed. Run: pip install comtypes")

        powerpoint = None
        presentation = None
        try:
            # Initialize COM
            comtypes.client.CoInitialize()
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1  # Must be visible for export

            # Give PowerPoint time to initialize
            time.sleep(0.5)

            # Open presentation with window (some environments need this)
            presentation = powerpoint.Presentations.Open(
                str(self.path.absolute()),
                ReadOnly=True,
                Untitled=False,
                WithWindow=True  # Changed to True for better compatibility
            )

            # Wait for presentation to fully load
            time.sleep(0.5)

            total = presentation.Slides.Count
            self.pages = []

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                # Export each slide as PNG
                for i in range(1, total + 1):
                    if progress_callback:
                        progress_callback(i, total)

                    slide = presentation.Slides(i)
                    img_path = tmpdir_path / f"slide_{i}.png"

                    # Export slide as image (width based on DPI)
                    # PowerPoint default is 96 DPI, we want higher
                    export_width = int(1920 * (full_dpi / 96))

                    # Try export with retries
                    for retry in range(3):
                        try:
                            slide.Export(str(img_path), "PNG", export_width)
                            time.sleep(0.1)  # Small delay for file write
                            if img_path.exists() and img_path.stat().st_size > 0:
                                break
                        except Exception:
                            time.sleep(0.2)

                    if img_path.exists() and img_path.stat().st_size > 0:
                        img = Image.open(img_path)
                        thumbnail = img.copy()
                        thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

                        page = Page(index=i - 1, thumbnail=thumbnail)
                        page.compute_phash()
                        self.pages.append(page)
                    else:
                        # Create placeholder if export failed
                        thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                        page = Page(index=i - 1, thumbnail=thumbnail)
                        page.compute_phash()
                        self.pages.append(page)

        finally:
            try:
                if presentation:
                    presentation.Close()
            except Exception:
                pass
            try:
                if powerpoint:
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

            page.full_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
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
