"""Main window for PySide6 GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Optional, List
import sys

from PySide6.QtCore import Qt, Signal, QMimeData, QPoint, QRectF, QTimer
from PySide6.QtGui import (
    QAction, QPainter, QPen, QColor, QBrush,
    QDragEnterEvent, QDropEvent, QPixmap, QFont
)
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QScrollArea, QLabel, QMenuBar, QMenu,
    QToolBar, QStatusBar, QFileDialog, QMessageBox,
    QProgressDialog, QApplication, QFrame, QGraphicsView,
    QGraphicsScene, QGraphicsPixmapItem, QGraphicsLineItem,
    QGraphicsRectItem
)

from ..core.document import Document
from ..core.page_matcher import PageMatcher, MatchingResult, MatchStatus
from ..core.exclusion_zone import ExclusionZoneSet
from ..core.session import Session
from ..utils.image_utils import pil_to_qpixmap


class PageThumbnail(QFrame):
    """Widget displaying a single page thumbnail."""

    clicked = Signal(int, str)  # page_index, side ("left" or "right")
    double_clicked = Signal(int, str)

    def __init__(
        self,
        page_index: int,
        side: str,
        parent: Optional[QWidget] = None
    ):
        super().__init__(parent)
        self.page_index = page_index
        self.side = side
        self._original_pixmap: Optional[QPixmap] = None
        self._selected = False
        self._match_status: Optional[MatchStatus] = None

        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        self.setLineWidth(2)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setSizePolicy(
            self.image_label.sizePolicy().horizontalPolicy(),
            self.image_label.sizePolicy().verticalPolicy()
        )

        self.page_label = QLabel(f"Page {page_index + 1}")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setPointSize(9)
        self.page_label.setFont(font)

        layout.addWidget(self.image_label, 1)
        layout.addWidget(self.page_label, 0)

        self.setAcceptDrops(True)
        self._update_style()

    def set_pixmap(self, pixmap: QPixmap) -> None:
        """Set the thumbnail image."""
        self._original_pixmap = pixmap
        self._update_scaled_pixmap()

    def _update_scaled_pixmap(self, target_width: Optional[int] = None) -> None:
        """Update the displayed pixmap scaled to target width."""
        if self._original_pixmap is None:
            return

        if target_width is None:
            # Get width from parent
            parent = self.parent()
            if parent:
                target_width = parent.width() - 50
            else:
                target_width = 300

        target_width = max(100, target_width if target_width else 300)

        # Scale maintaining aspect ratio
        orig_w = self._original_pixmap.width()
        orig_h = self._original_pixmap.height()
        scale = target_width / orig_w
        target_height = int(orig_h * scale)

        scaled = self._original_pixmap.scaled(
            target_width, target_height,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self.image_label.setPixmap(scaled)

        # Update size hints
        self.setMinimumSize(target_width + 20, target_height + 40)
        self.setMaximumHeight(target_height + 60)

    def set_selected(self, selected: bool) -> None:
        """Set selection state."""
        self._selected = selected
        self._update_style()

    def set_match_status(self, status: Optional[MatchStatus]) -> None:
        """Set match status for coloring."""
        self._match_status = status
        self._update_style()

    def _update_style(self) -> None:
        """Update visual style based on state."""
        if self._selected:
            self.setStyleSheet("""
                PageThumbnail {
                    background-color: #cce5ff;
                    border: 3px solid #004085;
                }
            """)
        elif self._match_status == MatchStatus.UNMATCHED_LEFT:
            self.setStyleSheet("""
                PageThumbnail {
                    background-color: #fff3cd;
                    border: 2px solid #856404;
                }
            """)
        elif self._match_status == MatchStatus.UNMATCHED_RIGHT:
            self.setStyleSheet("""
                PageThumbnail {
                    background-color: #d4edda;
                    border: 2px solid #155724;
                }
            """)
        else:
            self.setStyleSheet("""
                PageThumbnail {
                    background-color: white;
                    border: 1px solid #ccc;
                }
                PageThumbnail:hover {
                    border: 2px solid #007bff;
                }
            """)

    def mousePressEvent(self, event):
        """Handle mouse press."""
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.page_index, self.side)
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        """Handle double click."""
        if event.button() == Qt.MouseButton.LeftButton:
            self.double_clicked.emit(self.page_index, self.side)
        super().mouseDoubleClickEvent(event)

    def get_center_global(self) -> QPoint:
        """Get center position in global coordinates."""
        center = self.rect().center()
        return self.mapToGlobal(center)


class DocumentPanel(QScrollArea):
    """Scrollable panel containing document page thumbnails."""

    page_clicked = Signal(int, str)
    page_double_clicked = Signal(int, str)
    file_dropped = Signal(str)

    def __init__(self, side: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.side = side
        self.document: Optional[Document] = None
        self.thumbnails: List[PageThumbnail] = []
        self._selected_index: Optional[int] = None

        self.setWidgetResizable(True)
        self.setAcceptDrops(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # Container widget
        self.container = QWidget()
        self.layout = QVBoxLayout(self.container)
        self.layout.setSpacing(10)
        self.layout.setContentsMargins(0, 0, 10, 10)

        # Header for file info
        self.header_widget = QWidget()
        self.header_widget.setStyleSheet("""
            QWidget {
                background-color: #343a40;
            }
        """)
        header_layout = QVBoxLayout(self.header_widget)
        header_layout.setContentsMargins(10, 8, 10, 8)
        header_layout.setSpacing(2)

        self.file_name_label = QLabel(f"{'Left' if side == 'left' else 'Right'} Document")
        self.file_name_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 12px;
                font-weight: bold;
            }
        """)

        self.file_path_label = QLabel("No file loaded")
        self.file_path_label.setStyleSheet("""
            QLabel {
                color: #adb5bd;
                font-size: 9px;
            }
        """)

        header_layout.addWidget(self.file_name_label)
        header_layout.addWidget(self.file_path_label)

        self.layout.addWidget(self.header_widget)
        self.layout.addStretch()

        self.setWidget(self.container)

        # Placeholder
        self.placeholder = QLabel(
            f"Drop {'Left' if side == 'left' else 'Right'} document here\n"
            f"(PDF or PowerPoint)"
        )
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 14px;
                border: 2px dashed #ccc;
                border-radius: 10px;
                padding: 40px;
                background-color: #f8f9fa;
            }
        """)
        self.layout.insertWidget(1, self.placeholder)

        # Timer for delayed resize updates
        self._resize_timer = QTimer()
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._update_all_thumbnails)

    def resizeEvent(self, event) -> None:
        """Handle resize event."""
        super().resizeEvent(event)
        # Delay update to avoid excessive redraws
        self._resize_timer.start(100)

    def _update_all_thumbnails(self) -> None:
        """Update all thumbnail sizes based on current panel width."""
        target_width = self.width() - 50
        for thumb in self.thumbnails:
            if thumb._original_pixmap is not None:
                thumb._update_scaled_pixmap(target_width)

    def set_document(self, document: Document) -> None:
        """Set the document to display."""
        self.document = document
        self._update_header()
        self._rebuild_thumbnails()

    def _update_header(self) -> None:
        """Update header with file information."""
        if self.document is None:
            self.file_name_label.setText(f"{'Left' if self.side == 'left' else 'Right'} Document")
            self.file_path_label.setText("No file loaded")
        else:
            self.file_name_label.setText(self.document.name)
            # Truncate path if too long
            path_str = str(self.document.path.parent)
            if len(path_str) > 50:
                path_str = "..." + path_str[-47:]
            self.file_path_label.setText(f"{path_str}  ({self.document.page_count} pages)")

    def _rebuild_thumbnails(self) -> None:
        """Rebuild thumbnail widgets from document."""
        # Clear existing
        for thumb in self.thumbnails:
            thumb.deleteLater()
        self.thumbnails.clear()

        if self.document is None:
            self.placeholder.show()
            return

        self.placeholder.hide()

        # Create thumbnails
        for page in self.document.pages:
            thumb = PageThumbnail(page.index, self.side)
            thumb.clicked.connect(self._on_page_clicked)
            thumb.double_clicked.connect(self._on_page_double_clicked)

            if page.thumbnail:
                pixmap = pil_to_qpixmap(page.thumbnail)
                thumb.set_pixmap(pixmap)

            self.thumbnails.append(thumb)
            self.layout.insertWidget(self.layout.count() - 1, thumb)

    def _on_page_clicked(self, index: int, side: str):
        """Handle page click."""
        self.page_clicked.emit(index, side)

    def _on_page_double_clicked(self, index: int, side: str):
        """Handle page double click."""
        self.page_double_clicked.emit(index, side)

    def set_selected(self, index: Optional[int]) -> None:
        """Set selected page."""
        if self._selected_index is not None and self._selected_index < len(self.thumbnails):
            self.thumbnails[self._selected_index].set_selected(False)

        self._selected_index = index

        if index is not None and index < len(self.thumbnails):
            self.thumbnails[index].set_selected(True)

    def update_match_status(self, result: MatchingResult) -> None:
        """Update thumbnails with match status."""
        for thumb in self.thumbnails:
            if self.side == "left":
                match = result.get_match_for_left(thumb.page_index)
            else:
                match = result.get_match_for_right(thumb.page_index)

            if match:
                thumb.set_match_status(match.status)
            else:
                thumb.set_match_status(None)

    def get_thumbnail_position(self, index: int) -> Optional[QPoint]:
        """Get position of a thumbnail for drawing links."""
        if index < len(self.thumbnails):
            thumb = self.thumbnails[index]
            # Get center of thumbnail relative to main window
            center = thumb.rect().center()
            if self.side == "left":
                # Right edge
                pos = QPoint(thumb.width(), center.y())
            else:
                # Left edge
                pos = QPoint(0, center.y())
            return thumb.mapTo(self.window(), pos)
        return None

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        """Handle drag enter."""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                path = url.toLocalFile().lower()
                if path.endswith(('.pdf', '.pptx', '.ppt')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        """Handle file drop - supports multiple files."""
        urls = event.mimeData().urls()
        valid_files = []
        for url in urls:
            path = url.toLocalFile()
            if path.lower().endswith(('.pdf', '.pptx', '.ppt')):
                valid_files.append(path)

        if len(valid_files) >= 2:
            # Two files: emit for both left and right
            # Use parent window to load both
            window = self.window()
            if hasattr(window, '_load_document'):
                window._load_document(valid_files[0], "left")
                window._load_document(valid_files[1], "right")
        elif len(valid_files) == 1:
            # Single file: emit for this panel's side
            self.file_dropped.emit(valid_files[0])

        event.acceptProposedAction()


class LinkOverlay(QWidget):
    """Transparent overlay that draws lines between matched pages."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.links: List[tuple] = []  # [(left_pos, right_pos, status, similarity), ...]

    def set_links(self, links: List[tuple]) -> None:
        """Set links to draw."""
        self.links = links
        self.update()

    def paintEvent(self, event) -> None:
        """Draw the link lines."""
        if not self.links:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        for left_pos, right_pos, status, similarity in self.links:
            if left_pos is None or right_pos is None:
                continue

            # Choose color based on status and similarity
            if status == MatchStatus.MATCHED:
                if similarity >= 0.95:
                    color = QColor(40, 167, 69)  # Green - very similar
                elif similarity >= 0.8:
                    color = QColor(255, 193, 7)  # Yellow - some differences
                else:
                    color = QColor(220, 53, 69)  # Red - significant differences
            else:
                color = QColor(108, 117, 125)  # Gray - unmatched

            pen = QPen(color)
            pen.setWidth(2)
            painter.setPen(pen)

            # Draw bezier curve
            painter.drawLine(left_pos, right_pos)

        painter.end()


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT/PDF Comparator (PySide6)")
        self.setMinimumSize(600, 400)
        self.resize(1200, 800)

        self.session = Session()
        self.left_doc: Optional[Document] = None
        self.right_doc: Optional[Document] = None
        self.matching_result: Optional[MatchingResult] = None

        self._selected_left: Optional[int] = None
        self._selected_right: Optional[int] = None

        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()

        # Enable drag and drop on main window
        self.setAcceptDrops(True)

    def _setup_ui(self) -> None:
        """Set up the main UI layout."""
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Create splitter for resizable panels
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left document panel
        self.left_panel = DocumentPanel("left")
        self.left_panel.page_clicked.connect(self._on_page_clicked)
        self.left_panel.page_double_clicked.connect(self._on_page_double_clicked)
        self.left_panel.file_dropped.connect(self._on_left_file_dropped)
        self.left_panel.setMinimumWidth(200)

        # Right document panel
        self.right_panel = DocumentPanel("right")
        self.right_panel.page_clicked.connect(self._on_page_clicked)
        self.right_panel.page_double_clicked.connect(self._on_page_double_clicked)
        self.right_panel.file_dropped.connect(self._on_right_file_dropped)
        self.right_panel.setMinimumWidth(200)

        self.splitter.addWidget(self.left_panel)
        self.splitter.addWidget(self.right_panel)
        self.splitter.setSizes([600, 600])

        main_layout.addWidget(self.splitter)

        # Link overlay (positioned over entire window)
        self.link_overlay = LinkOverlay(central)

        # Update overlay on scroll
        self.left_panel.verticalScrollBar().valueChanged.connect(self._update_links)
        self.right_panel.verticalScrollBar().valueChanged.connect(self._update_links)

        # Timer for delayed link updates
        self.link_update_timer = QTimer()
        self.link_update_timer.setSingleShot(True)
        self.link_update_timer.timeout.connect(self._do_update_links)

    def _setup_menu(self) -> None:
        """Set up the menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("File")

        open_left = QAction("Open Left...", self)
        open_left.setShortcut("Ctrl+O")
        open_left.triggered.connect(self._open_left_file)
        file_menu.addAction(open_left)

        open_right = QAction("Open Right...", self)
        open_right.setShortcut("Ctrl+Shift+O")
        open_right.triggered.connect(self._open_right_file)
        file_menu.addAction(open_right)

        file_menu.addSeparator()

        save_session = QAction("Save Session...", self)
        save_session.setShortcut("Ctrl+S")
        save_session.triggered.connect(self._save_session)
        file_menu.addAction(save_session)

        load_session = QAction("Load Session...", self)
        load_session.setShortcut("Ctrl+L")
        load_session.triggered.connect(self._load_session)
        file_menu.addAction(load_session)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Compare menu
        compare_menu = menubar.addMenu("Compare")

        run_compare = QAction("Run Comparison", self)
        run_compare.setShortcut("F5")
        run_compare.triggered.connect(self._run_comparison)
        compare_menu.addAction(run_compare)

        compare_menu.addSeparator()

        clear_links = QAction("Clear Manual Links", self)
        clear_links.triggered.connect(self._clear_manual_links)
        compare_menu.addAction(clear_links)

        # Export menu
        export_menu = menubar.addMenu("Export")

        export_pdf = QAction("Export to PDF...", self)
        export_pdf.triggered.connect(self._export_pdf)
        export_menu.addAction(export_pdf)

        export_html = QAction("Export to HTML...", self)
        export_html.triggered.connect(self._export_html)
        export_menu.addAction(export_html)

    def _setup_toolbar(self) -> None:
        """Set up the toolbar."""
        toolbar = QToolBar()
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        open_left_btn = QAction("Open Left", self)
        open_left_btn.triggered.connect(self._open_left_file)
        toolbar.addAction(open_left_btn)

        open_right_btn = QAction("Open Right", self)
        open_right_btn.triggered.connect(self._open_right_file)
        toolbar.addAction(open_right_btn)

        toolbar.addSeparator()

        compare_btn = QAction("Compare", self)
        compare_btn.triggered.connect(self._run_comparison)
        toolbar.addAction(compare_btn)

        toolbar.addSeparator()

        export_pdf_btn = QAction("Export PDF", self)
        export_pdf_btn.triggered.connect(self._export_pdf)
        toolbar.addAction(export_pdf_btn)

        export_html_btn = QAction("Export HTML", self)
        export_html_btn.triggered.connect(self._export_html)
        toolbar.addAction(export_html_btn)

    def _setup_statusbar(self) -> None:
        """Set up the status bar."""
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.statusbar.showMessage("Ready - Drop files to begin")

    def _open_file_dialog(self) -> Optional[str]:
        """Open file dialog and return selected path."""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Document",
            "",
            "Documents (*.pdf *.pptx *.ppt);;PDF Files (*.pdf);;PowerPoint Files (*.pptx *.ppt)"
        )
        return path if path else None

    def _open_left_file(self) -> None:
        """Open file dialog for left document."""
        path = self._open_file_dialog()
        if path:
            self._load_document(path, "left")

    def _open_right_file(self) -> None:
        """Open file dialog for right document."""
        path = self._open_file_dialog()
        if path:
            self._load_document(path, "right")

    def _on_left_file_dropped(self, path: str) -> None:
        """Handle file dropped on left panel."""
        self._load_document(path, "left")

    def _on_right_file_dropped(self, path: str) -> None:
        """Handle file dropped on right panel."""
        self._load_document(path, "right")

    def _load_document(self, path: str, side: str) -> None:
        """Load a document file."""
        try:
            doc = Document.from_file(path)

            # Show progress dialog
            progress = QProgressDialog(
                f"Loading {Path(path).name}...",
                "Cancel", 0, 100, self
            )
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.show()

            def progress_callback(current, total):
                progress.setValue(int(current / total * 100))
                QApplication.processEvents()

            doc.load(progress_callback=progress_callback)
            progress.close()

            if side == "left":
                self.left_doc = doc
                self.left_panel.set_document(doc)
                self.session.left_document_path = path
            else:
                self.right_doc = doc
                self.right_panel.set_document(doc)
                self.session.right_document_path = path

            self._update_status()
            self._update_links()

        except Exception as e:
            QMessageBox.critical(
                self, "Error",
                f"Failed to load document:\n{str(e)}"
            )

    def _run_comparison(self) -> None:
        """Run page matching comparison."""
        if not self.left_doc or not self.right_doc:
            QMessageBox.warning(
                self, "Warning",
                "Please load both documents before comparing."
            )
            return

        # Show progress
        progress = QProgressDialog(
            "Matching pages...", "Cancel", 0, 100, self
        )
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        def progress_callback(current, total, msg=""):
            progress.setLabelText(msg or "Matching pages...")
            progress.setValue(int(current / total * 100) if total > 0 else 0)
            QApplication.processEvents()

        matcher = PageMatcher()
        self.matching_result = matcher.match(
            self.left_doc, self.right_doc,
            progress_callback=progress_callback
        )
        progress.close()

        self.session.matching_result = self.matching_result

        # Update UI
        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_links()
        self._update_status()

    def _on_page_clicked(self, index: int, side: str) -> None:
        """Handle page click for manual linking."""
        if side == "left":
            self._selected_left = index
            self.left_panel.set_selected(index)
        else:
            self._selected_right = index
            self.right_panel.set_selected(index)

        # If both selected, create manual link
        if self._selected_left is not None and self._selected_right is not None:
            self._create_manual_link(self._selected_left, self._selected_right)
            self._selected_left = None
            self._selected_right = None
            self.left_panel.set_selected(None)
            self.right_panel.set_selected(None)

    def _on_page_double_clicked(self, index: int, side: str) -> None:
        """Handle page double click for viewing details."""
        # TODO: Open detail view with diff highlighting
        pass

    def _create_manual_link(self, left_index: int, right_index: int) -> None:
        """Create a manual link between pages."""
        if self.matching_result is None:
            self.matching_result = MatchingResult()

        self.matching_result.set_manual_match(left_index, right_index)
        self.session.matching_result = self.matching_result

        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_links()

        self.statusbar.showMessage(
            f"Linked: Left page {left_index + 1} ↔ Right page {right_index + 1}"
        )

    def _clear_manual_links(self) -> None:
        """Clear all manual links."""
        if self.matching_result:
            # Re-run automatic matching
            self._run_comparison()

    def _update_links(self) -> None:
        """Schedule link overlay update."""
        self.link_update_timer.start(50)

    def _do_update_links(self) -> None:
        """Actually update the link overlay."""
        if self.matching_result is None:
            self.link_overlay.set_links([])
            return

        links = []
        for match in self.matching_result.matches:
            if match.status == MatchStatus.MATCHED:
                left_pos = self.left_panel.get_thumbnail_position(match.left_index)
                right_pos = self.right_panel.get_thumbnail_position(match.right_index)
                if left_pos and right_pos:
                    links.append((left_pos, right_pos, match.status, match.similarity))

        self.link_overlay.set_links(links)
        self.link_overlay.setGeometry(self.centralWidget().rect())
        self.link_overlay.raise_()

    def _update_status(self) -> None:
        """Update status bar."""
        parts = []
        if self.left_doc:
            parts.append(f"Left: {self.left_doc.name} ({self.left_doc.page_count} pages)")
        if self.right_doc:
            parts.append(f"Right: {self.right_doc.name} ({self.right_doc.page_count} pages)")
        if self.matching_result:
            matched = sum(1 for m in self.matching_result.matches if m.status == MatchStatus.MATCHED)
            parts.append(f"Matched: {matched}")

        self.statusbar.showMessage(" | ".join(parts) if parts else "Ready")

    def _save_session(self) -> None:
        """Save current session."""
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Session",
            "", "Session Files (*.json)"
        )
        if path:
            self.session.save(path)
            self.statusbar.showMessage(f"Session saved to {path}")

    def _load_session(self) -> None:
        """Load a session file."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Session",
            "", "Session Files (*.json)"
        )
        if path:
            try:
                self.session = Session.load(path)
                if self.session.left_document_path:
                    self._load_document(self.session.left_document_path, "left")
                if self.session.right_document_path:
                    self._load_document(self.session.right_document_path, "right")
                self.matching_result = self.session.matching_result
                if self.matching_result:
                    self.left_panel.update_match_status(self.matching_result)
                    self.right_panel.update_match_status(self.matching_result)
                    self._update_links()
                self.statusbar.showMessage(f"Session loaded from {path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load session:\n{e}")

    def _export_pdf(self) -> None:
        """Export comparison report to PDF."""
        QMessageBox.information(
            self, "Export PDF",
            "PDF export will be implemented in the full version."
        )

    def _export_html(self) -> None:
        """Export comparison report to HTML."""
        QMessageBox.information(
            self, "Export HTML",
            "HTML export will be implemented in the full version."
        )

    def resizeEvent(self, event) -> None:
        """Handle window resize."""
        super().resizeEvent(event)
        self._update_links()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        """Handle drag enter on main window."""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid_files = []
            for url in urls:
                path = url.toLocalFile().lower()
                if path.endswith(('.pdf', '.pptx', '.ppt')):
                    valid_files.append(url.toLocalFile())
            if valid_files:
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        """Handle file drop on main window."""
        urls = event.mimeData().urls()
        valid_files = []
        for url in urls:
            path = url.toLocalFile()
            if path.lower().endswith(('.pdf', '.pptx', '.ppt')):
                valid_files.append(path)

        if len(valid_files) >= 2:
            # Two files dropped - load as left and right
            self._load_document(valid_files[0], "left")
            self._load_document(valid_files[1], "right")
        elif len(valid_files) == 1:
            # Single file - load to empty side or left
            if self.left_doc is None:
                self._load_document(valid_files[0], "left")
            elif self.right_doc is None:
                self._load_document(valid_files[0], "right")
            else:
                self._load_document(valid_files[0], "left")

        event.acceptProposedAction()
