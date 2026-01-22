"""Document abstraction for PDF and PowerPoint files."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Optional, List, Tuple, Dict, Any
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

# Last PowerPoint error for diagnostics
_last_ppt_error: Optional[str] = None


def get_last_ppt_error() -> Optional[str]:
    """Get the last PowerPoint conversion error message."""
    return _last_ppt_error


def diagnose_powerpoint() -> dict:
    """Diagnose PowerPoint availability and COM setup.

    Returns:
        dict with diagnostic information
    """
    import sys
    result = {
        'platform': sys.platform,
        'python_bits': 64 if sys.maxsize > 2**32 else 32,
        'powerpoint_available': False,
        'powerpoint_version': None,
        'comtypes_installed': False,
        'error': None
    }

    if sys.platform != 'win32':
        result['error'] = 'Not running on Windows'
        return result

    try:
        import comtypes
        import comtypes.client
        result['comtypes_installed'] = True
    except ImportError:
        result['error'] = 'comtypes not installed'
        return result

    try:
        comtypes.CoInitialize()
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        result['powerpoint_available'] = True
        try:
            result['powerpoint_version'] = powerpoint.Version
        except Exception:
            result['powerpoint_version'] = 'Unknown'

        # Try to quit the test instance
        try:
            powerpoint.Quit()
        except Exception:
            pass

    except Exception as e:
        result['error'] = str(e)

    return result

# Disk cache directory
_CACHE_DIR: Optional[Path] = None


def _get_cache_dir() -> Path:
    """Get or create the cache directory."""
    global _CACHE_DIR
    if _CACHE_DIR is None:
        # Use user's home directory for cache
        home = Path.home()
        _CACHE_DIR = home / ".pptpdf_cache"
        _CACHE_DIR.mkdir(exist_ok=True)
    return _CACHE_DIR


def _get_cache_key(file_path: Path) -> str:
    """Generate a cache key based on file path and modification time."""
    mtime = file_path.stat().st_mtime
    size = file_path.stat().st_size
    key_str = f"{file_path.absolute()}|{mtime}|{size}"
    return hashlib.md5(key_str.encode()).hexdigest()


def _get_cached_thumbnails(file_path: Path) -> Optional[Dict[int, Path]]:
    """Check if cached thumbnails exist for this file.

    Returns:
        Dict mapping page index to cached PNG path, or None if not cached
    """
    cache_key = _get_cache_key(file_path)
    cache_subdir = _get_cache_dir() / cache_key

    if not cache_subdir.exists():
        return None

    meta_file = cache_subdir / "meta.json"
    if not meta_file.exists():
        return None

    try:
        with open(meta_file, 'r') as f:
            meta = json.load(f)

        # Verify the cache is still valid
        if meta.get('source_path') != str(file_path.absolute()):
            return None

        # Find all cached pages
        cached = {}
        for png in cache_subdir.glob("page_*.png"):
            try:
                idx = int(png.stem.split('_')[1])
                cached[idx] = png
            except (ValueError, IndexError):
                pass

        if len(cached) >= meta.get('page_count', 0):
            return cached

    except Exception:
        pass

    return None


def _save_to_cache(file_path: Path, pages: List['Page']) -> None:
    """Save thumbnails to disk cache."""
    cache_key = _get_cache_key(file_path)
    cache_subdir = _get_cache_dir() / cache_key
    cache_subdir.mkdir(exist_ok=True)

    # Save metadata
    meta = {
        'source_path': str(file_path.absolute()),
        'page_count': len(pages),
        'cached_at': time.time()
    }
    with open(cache_subdir / "meta.json", 'w') as f:
        json.dump(meta, f)

    # Save thumbnails
    for page in pages:
        if page.thumbnail is not None:
            png_path = cache_subdir / f"page_{page.index}.png"
            page.thumbnail.save(png_path, "PNG")


def clear_cache() -> None:
    """Clear all cached thumbnails."""
    import shutil
    cache_dir = _get_cache_dir()
    if cache_dir.exists():
        shutil.rmtree(cache_dir)
        cache_dir.mkdir(exist_ok=True)


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
        # Keep PowerPoint hidden (WindowState: 2=Minimized, Visible=0 may cause export issues)
        try:
            powerpoint.Visible = 1  # Must be visible for export to work
            powerpoint.WindowState = 2  # ppWindowMinimized = 2
        except Exception:
            powerpoint.Visible = 1  # Fallback to visible
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
    """Represents a document (PDF or PowerPoint) as a collection of pages.

    Optimizations:
    - Disk cache: Thumbnails are cached to disk for instant reload
    - Deferred pHash: pHash is computed only when needed for comparison
    """

    path: Path
    doc_type: DocumentType
    pages: List[Page] = field(default_factory=list)
    _loaded: bool = False
    _hashes_computed: bool = False
    _thumbnail_size: tuple[int, int] = field(default=(1200, 900), repr=False)

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
        thumbnail_size: tuple[int, int] = (1200, 900),  # 高画質サムネイル
        full_dpi: int = 200,
        progress_callback: Optional[callable] = None,
        use_cache: bool = True
    ) -> None:
        """Load all pages from the document.

        Args:
            thumbnail_size: Size for thumbnail images (width, height)
            full_dpi: DPI for full resolution images
            progress_callback: Optional callback(current, total) for progress
            use_cache: Whether to use disk cache (default True)
        """
        if self._loaded:
            return

        self._thumbnail_size = thumbnail_size

        # Try loading from cache first
        if use_cache:
            cached = _get_cached_thumbnails(self.path)
            if cached:
                print(f"[DEBUG] Loading from cache: {len(cached)} pages")
                self._load_from_cache(cached, progress_callback)
                self._loaded = True
                return

        # Load from source
        if self.doc_type == DocumentType.PDF:
            self._load_pdf(thumbnail_size, full_dpi, progress_callback)
        elif self.doc_type in (DocumentType.PPTX, DocumentType.PPT):
            self._load_pptx(thumbnail_size, full_dpi, progress_callback)

        self._loaded = True

        # Save to cache in background (non-blocking)
        if use_cache and self.pages:
            import threading

            def save_cache_background():
                try:
                    _save_to_cache(self.path, self.pages)
                    print(f"[DEBUG] Saved {len(self.pages)} pages to cache (background)")
                except Exception as e:
                    print(f"[DEBUG] Cache save failed: {e}")

            cache_thread = threading.Thread(target=save_cache_background, daemon=True)
            cache_thread.start()

    def _load_from_cache(
        self,
        cached: Dict[int, Path],
        progress_callback: Optional[callable]
    ) -> None:
        """Load thumbnails from disk cache (parallelized)."""
        total = len(cached)

        def load_single_page(args):
            """Worker function to load a single cached page."""
            idx, cache_path = args
            try:
                thumbnail = Image.open(cache_path)
                return idx, thumbnail
            except Exception:
                return idx, None

        # Parallel cache loading
        import os
        max_workers = min(8, max(4, os.cpu_count() or 4))

        work_items = [(i, cached[i]) for i in range(total) if i in cached]
        results = {}

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(load_single_page, item): item[0] for item in work_items}

            completed = 0
            for future in as_completed(futures):
                completed += 1
                if progress_callback:
                    progress_callback(completed, total)

                try:
                    idx, thumbnail = future.result()
                    results[idx] = thumbnail
                except Exception:
                    pass

        # Build pages list in order
        self.pages = []
        for i in range(total):
            if i in results and results[i] is not None:
                page = Page(index=i, thumbnail=results[i])
            else:
                # Missing page - create placeholder
                thumbnail = Image.new('RGB', self._thumbnail_size, color=(200, 200, 200))
                page = Page(index=i, thumbnail=thumbnail)
            self.pages.append(page)

    def ensure_hashes_computed(self, progress_callback: Optional[callable] = None) -> None:
        """Compute pHashes for all pages if not already done.

        This is called before comparison to ensure hashes are available.
        Deferred computation saves time during initial load.
        """
        if self._hashes_computed:
            return

        # Check which pages need hashes
        pages_needing_hash = [p for p in self.pages if p.phash is None and p.thumbnail is not None]

        if not pages_needing_hash:
            self._hashes_computed = True
            return

        print(f"[DEBUG] Computing pHashes for {len(pages_needing_hash)} pages...")

        # Compute in parallel
        _compute_phashes_parallel(self.pages, hash_size=16, max_workers=4)
        self._hashes_computed = True

    def _load_pdf(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Load pages from PDF file using PyMuPDF (parallelized).

        Note: pHash computation is deferred until ensure_hashes_computed() is called.
        """
        import fitz  # PyMuPDF
        import os

        doc = fitz.open(str(self.path))
        total = len(doc)

        # Calculate zoom for thumbnail (72 DPI base, we want ~216 DPI equivalent)
        zoom = 3.0  # 3x zoom for better quality

        def render_page(page_index: int) -> Tuple[int, Optional[Image.Image]]:
            """Render a single PDF page."""
            try:
                # Each thread needs its own document handle for thread safety
                thread_doc = fitz.open(str(self.path))
                page = thread_doc.load_page(page_index)
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)

                # Convert to PIL Image
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                # Create thumbnail
                thumbnail = img.copy()
                thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)

                thread_doc.close()
                return page_index, thumbnail
            except Exception as e:
                print(f"[DEBUG] Error rendering PDF page {page_index}: {e}")
                return page_index, None

        # Use ThreadPoolExecutor for parallel rendering
        max_workers = min(8, max(4, os.cpu_count() or 4))
        results = {}

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(render_page, i): i for i in range(total)}

            completed = 0
            for future in as_completed(futures):
                completed += 1
                if progress_callback:
                    progress_callback(completed, total)

                try:
                    page_index, thumbnail = future.result()
                    results[page_index] = thumbnail
                except Exception as e:
                    print(f"[DEBUG] PDF page future error: {e}")

        # Build pages list in order
        self.pages = []
        for i in range(total):
            if i in results and results[i] is not None:
                page_obj = Page(index=i, thumbnail=results[i])
            else:
                # Create placeholder for failed pages
                placeholder = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                page_obj = Page(index=i, thumbnail=placeholder)
            self.pages.append(page_obj)

        doc.close()
        # Note: pHash computation deferred to ensure_hashes_computed()
        # Note: Full images are loaded on-demand to save memory

    def _load_pptx(
        self,
        thumbnail_size: tuple[int, int],
        full_dpi: int,
        progress_callback: Optional[callable]
    ) -> None:
        """Load pages from PowerPoint file."""
        global _last_ppt_error
        from pptx import Presentation

        prs = Presentation(str(self.path))
        total = len(prs.slides)

        # Try different conversion methods in order of preference
        conversion_methods = [
            ("PowerPoint COM", self._convert_pptx_via_com),
            ("LibreOffice", self._convert_pptx_via_pdf),
        ]

        errors = []
        for method_name, method in conversion_methods:
            try:
                print(f"[DEBUG] Trying PPTX conversion via {method_name}...")
                method(thumbnail_size, full_dpi, progress_callback)
                print(f"[DEBUG] PPTX conversion via {method_name} succeeded")
                _last_ppt_error = None  # Clear error on success
                return
            except Exception as e:
                error_msg = f"{method_name}: {e}"
                print(f"[DEBUG] {error_msg}")
                errors.append(error_msg)
                continue

        # Fallback: create placeholder pages
        _last_ppt_error = "; ".join(errors)
        print(f"[DEBUG] All conversion methods failed: {_last_ppt_error}")
        self.pages = []
        for i in range(total):
            if progress_callback:
                progress_callback(i + 1, total)

            # Create a placeholder thumbnail with text
            thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
            page = Page(index=i, thumbnail=thumbnail)
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
        - BULK EXPORT: Export all slides at once (5-10x faster than individual)
        - Parallel image loading after export
        - Polling instead of fixed sleep
        """
        import sys
        import os
        if sys.platform != 'win32':
            raise RuntimeError("PowerPoint COM is only available on Windows")

        try:
            import comtypes
            import comtypes.client
        except ImportError:
            raise RuntimeError("comtypes not installed. Run: pip install comtypes")

        # Diagnose Python/PowerPoint architecture
        python_bits = 64 if sys.maxsize > 2**32 else 32
        print(f"[DEBUG] Python architecture: {python_bits}-bit")

        presentation = None
        powerpoint = None
        should_close_ppt = True

        try:
            # Initialize COM
            try:
                comtypes.CoInitialize()
            except Exception as e:
                raise RuntimeError(f"COM initialization failed: {e}")

            # Use cached PowerPoint instance
            try:
                powerpoint, should_close_ppt = get_powerpoint_instance()
            except Exception as e:
                raise RuntimeError(f"Failed to create PowerPoint instance (Python {python_bits}-bit): {e}")

            # Open presentation without window (minimized PowerPoint)
            presentation = powerpoint.Presentations.Open(
                str(self.path.absolute()),
                ReadOnly=True,
                Untitled=False,
                WithWindow=False  # No window - faster and less intrusive
            )

            total = presentation.Slides.Count

            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                # BULK EXPORT: Export all slides at once (MUCH faster)
                # ppSaveAsPNG = 18
                if progress_callback:
                    progress_callback(1, total + 1, "Exporting all slides...")

                export_path = tmpdir_path / "slides"
                export_path.mkdir(exist_ok=True)

                # Export resolution (width in pixels)
                export_width = 2560

                try:
                    # Bulk export - PowerPoint creates Slide1.PNG, Slide2.PNG, etc.
                    presentation.SaveAs(
                        str(export_path / "Slide"),
                        18,  # ppSaveAsPNG
                        False  # EmbedTrueTypeFonts
                    )
                except Exception as e:
                    print(f"[DEBUG] Bulk export failed: {e}, falling back to individual export")
                    # Fallback to individual export
                    self._convert_pptx_via_com_individual(
                        presentation, thumbnail_size, total, tmpdir_path, progress_callback
                    )
                    return

                # Wait for all files to be created
                png_files = _wait_for_files(export_path, total, timeout=30.0)

                if len(png_files) < total:
                    # Some files missing, try individual export as fallback
                    print(f"[DEBUG] Only {len(png_files)}/{total} files found, falling back")
                    self._convert_pptx_via_com_individual(
                        presentation, thumbnail_size, total, tmpdir_path, progress_callback
                    )
                    return

                # Parallel image loading
                def load_slide_image(args):
                    """Load and resize a single slide image."""
                    idx, img_path = args
                    try:
                        img = Image.open(img_path)
                        thumbnail = img.copy()
                        thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)
                        return idx, thumbnail
                    except Exception as e:
                        print(f"[DEBUG] Error loading slide {idx}: {e}")
                        return idx, None

                # Map PNG files to slide indices
                slide_files = {}
                for png_file in png_files:
                    # Extract slide number from filename (Slide1.PNG, Slide2.PNG, etc.)
                    match = re.search(r'(\d+)', png_file.stem)
                    if match:
                        slide_num = int(match.group(1))
                        slide_files[slide_num - 1] = png_file  # 0-indexed

                # Parallel loading
                max_workers = min(8, max(4, os.cpu_count() or 4))
                work_items = [(idx, path) for idx, path in slide_files.items()]
                results = {}

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {executor.submit(load_slide_image, item): item[0] for item in work_items}

                    completed = 0
                    for future in as_completed(futures):
                        completed += 1
                        if progress_callback:
                            progress_callback(completed + 1, total + 1)

                        try:
                            idx, thumbnail = future.result()
                            results[idx] = thumbnail
                        except Exception:
                            pass

                # Build pages list in order
                self.pages = []
                for i in range(total):
                    if i in results and results[i] is not None:
                        page = Page(index=i, thumbnail=results[i])
                    else:
                        # Create placeholder for missing slides
                        placeholder = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                        page = Page(index=i, thumbnail=placeholder)
                    self.pages.append(page)

                # pHash deferred to ensure_hashes_computed()

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

    def _convert_pptx_via_com_individual(
        self,
        presentation,
        thumbnail_size: tuple[int, int],
        total: int,
        tmpdir_path: Path,
        progress_callback: Optional[callable]
    ) -> None:
        """Fallback: Convert slides individually (slower but more reliable)."""
        self.pages = []
        export_width = 2560

        for i in range(1, total + 1):
            if progress_callback:
                progress_callback(i, total)

            slide = presentation.Slides(i)
            img_path = tmpdir_path / f"slide_{i}.png"

            try:
                slide.Export(str(img_path), "PNG", export_width)

                if _wait_for_file(img_path, timeout=3.0, min_size=1000):
                    img = Image.open(img_path)
                    thumbnail = img.copy()
                    thumbnail.thumbnail(thumbnail_size, Image.Resampling.LANCZOS)
                    page = Page(index=i - 1, thumbnail=thumbnail)
                    self.pages.append(page)
                else:
                    thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                    page = Page(index=i - 1, thumbnail=thumbnail)
                    self.pages.append(page)
            except Exception:
                thumbnail = Image.new('RGB', thumbnail_size, color=(200, 200, 200))
                page = Page(index=i - 1, thumbnail=thumbnail)
                self.pages.append(page)

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
