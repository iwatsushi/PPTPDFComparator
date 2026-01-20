"""Main window for PySide6 GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Optional, List
import sys

from PySide6.QtCore import Qt, Signal, QMimeData, QPoint, QRectF, QTimer
from PySide6.QtGui import (
    QAction, QPainter, QPen, QColor, QBrush,
    QDragEnterEvent, QDropEvent, QPixmap, QFont,
    QShortcut, QKeySequence
)
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QScrollArea, QLabel, QMenuBar, QMenu,
    QToolBar, QStatusBar, QFileDialog, QMessageBox,
    QProgressDialog, QApplication, QFrame, QGraphicsView,
    QGraphicsScene, QGraphicsPixmapItem, QGraphicsLineItem,
    QGraphicsRectItem, QSlider, QCheckBox, QDialog,
    QDialogButtonBox, QListWidget, QListWidgetItem,
    QPushButton, QSpinBox, QLineEdit, QComboBox,
    QFormLayout, QGroupBox, QColorDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView
)

try:
    from src.core.document import Document
    from src.core.page_matcher import PageMatcher, MatchingResult, MatchStatus
    from src.core.image_comparator import ImageComparator, DiffResult
    from src.core.exclusion_zone import ExclusionZone, ExclusionZoneSet, AppliesTo
    from src.core.session import Session
    from src.utils.image_utils import pil_to_qpixmap
except ImportError:
    from core.document import Document
    from core.page_matcher import PageMatcher, MatchingResult, MatchStatus
    from core.image_comparator import ImageComparator, DiffResult
    from core.exclusion_zone import ExclusionZone, ExclusionZoneSet, AppliesTo
    from core.session import Session
    from utils.image_utils import pil_to_qpixmap


class PageThumbnail(QFrame):
    """Widget displaying a single page thumbnail."""

    clicked = Signal(int, str)  # page_index, side ("left" or "right")
    double_clicked = Signal(int, str)
    exclusion_zone_drawn = Signal(str, int, float, float, float, float)  # side, page_idx, x, y, w, h

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
        self._highlight_pixmap: Optional[QPixmap] = None
        self._show_diff: bool = True
        self._selected = False
        self._match_status: Optional[MatchStatus] = None
        self._diff_result: Optional[DiffResult] = None

        # Exclusion zone drawing state
        self._drawing_mode: bool = False
        self._draw_start: Optional[QPoint] = None
        self._draw_current: Optional[QPoint] = None
        self._exclusion_zones: List[tuple] = []  # List of (x, y, w, h) in normalized coords

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
        self.image_label.setMouseTracking(True)

        self.page_label = QLabel(f"Page {page_index + 1}")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setPointSize(9)
        self.page_label.setFont(font)

        layout.addWidget(self.image_label, 1)
        layout.addWidget(self.page_label, 0)

        self.setAcceptDrops(True)
        self._update_style()

    def set_drawing_mode(self, enabled: bool) -> None:
        """Enable or disable exclusion zone drawing mode."""
        self._drawing_mode = enabled
        if enabled:
            self.setCursor(Qt.CursorShape.CrossCursor)
        else:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        self._draw_start = None
        self._draw_current = None
        self.update()

    def set_exclusion_zones(self, zones: List[tuple]) -> None:
        """Set exclusion zones to display as overlays."""
        self._exclusion_zones = zones
        self.update()

    def set_pixmap(self, pixmap: QPixmap) -> None:
        """Set the thumbnail image."""
        self._original_pixmap = pixmap
        self._highlight_pixmap = None
        self._update_scaled_pixmap()

    def set_diff_result(self, diff_result: Optional[DiffResult], highlight_pixmap: Optional[QPixmap] = None) -> None:
        """Set the diff result and highlighted image."""
        self._diff_result = diff_result
        self._highlight_pixmap = highlight_pixmap
        self._update_scaled_pixmap()

    def _update_scaled_pixmap(self, target_width: Optional[int] = None) -> None:
        """Update the displayed pixmap scaled to target width."""
        # Choose which pixmap to display
        pixmap_to_use = self._original_pixmap
        if self._show_diff and self._highlight_pixmap is not None:
            pixmap_to_use = self._highlight_pixmap

        if pixmap_to_use is None:
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
        orig_w = pixmap_to_use.width()
        orig_h = pixmap_to_use.height()
        scale = target_width / orig_w
        target_height = int(orig_h * scale)

        scaled = pixmap_to_use.scaled(
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
        elif self._match_status == MatchStatus.UNMATCHED_LEFT or self._match_status == MatchStatus.UNMATCHED_RIGHT:
            # Orange border for unmatched slides (only on one side)
            self.setStyleSheet("""
                PageThumbnail {
                    background-color: #ffe0b2;
                    border: 3px solid #ff9800;
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
            if self._drawing_mode:
                # Start drawing exclusion zone
                self._draw_start = event.position().toPoint()
                self._draw_current = self._draw_start
            else:
                self.clicked.emit(self.page_index, self.side)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """Handle mouse move."""
        if self._drawing_mode and self._draw_start is not None:
            self._draw_current = event.position().toPoint()
            self.update()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        """Handle mouse release."""
        if event.button() == Qt.MouseButton.LeftButton:
            if self._drawing_mode and self._draw_start is not None and self._draw_current is not None:
                # Calculate normalized coordinates relative to image
                pixmap = self.image_label.pixmap()
                if pixmap:
                    img_rect = self.image_label.geometry()
                    img_w = pixmap.width()
                    img_h = pixmap.height()

                    # Get image offset within label (centered)
                    offset_x = (img_rect.width() - img_w) // 2
                    offset_y = (img_rect.height() - img_h) // 2

                    # Adjust coordinates relative to image
                    x1 = self._draw_start.x() - img_rect.x() - offset_x
                    y1 = self._draw_start.y() - img_rect.y() - offset_y
                    x2 = self._draw_current.x() - img_rect.x() - offset_x
                    y2 = self._draw_current.y() - img_rect.y() - offset_y

                    # Normalize to 0-1 range
                    if img_w > 0 and img_h > 0:
                        nx = max(0, min(x1, x2)) / img_w
                        ny = max(0, min(y1, y2)) / img_h
                        nw = abs(x2 - x1) / img_w
                        nh = abs(y2 - y1) / img_h

                        # Only emit if zone is large enough
                        if nw > 0.01 and nh > 0.01:
                            self.exclusion_zone_drawn.emit(self.side, self.page_index, nx, ny, nw, nh)

                self._draw_start = None
                self._draw_current = None
                self.update()
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        """Handle double click."""
        if event.button() == Qt.MouseButton.LeftButton:
            self.double_clicked.emit(self.page_index, self.side)
        super().mouseDoubleClickEvent(event)

    def paintEvent(self, event):
        """Paint the widget with exclusion zone overlays."""
        super().paintEvent(event)

        pixmap = self.image_label.pixmap()
        if not pixmap:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        img_rect = self.image_label.geometry()
        img_w = pixmap.width()
        img_h = pixmap.height()

        # Get image offset within label (centered)
        offset_x = img_rect.x() + (img_rect.width() - img_w) // 2
        offset_y = img_rect.y() + (img_rect.height() - img_h) // 2

        # Draw existing exclusion zones
        painter.setBrush(QBrush(QColor(255, 0, 0, 50)))
        painter.setPen(QPen(QColor(255, 0, 0), 2, Qt.PenStyle.DotLine))
        for x, y, w, h in self._exclusion_zones:
            rect_x = offset_x + int(x * img_w)
            rect_y = offset_y + int(y * img_h)
            rect_w = int(w * img_w)
            rect_h = int(h * img_h)
            painter.drawRect(rect_x, rect_y, rect_w, rect_h)

        # Draw current selection rectangle
        if self._drawing_mode and self._draw_start and self._draw_current:
            painter.setBrush(QBrush(QColor(0, 120, 215, 50)))
            painter.setPen(QPen(QColor(0, 120, 215), 2))
            x1, y1 = self._draw_start.x(), self._draw_start.y()
            x2, y2 = self._draw_current.x(), self._draw_current.y()
            painter.drawRect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))

        painter.end()

    def get_center_global(self) -> QPoint:
        """Get center position in global coordinates."""
        center = self.rect().center()
        return self.mapToGlobal(center)


class DocumentPanel(QScrollArea):
    """Scrollable panel containing document page thumbnails."""

    page_clicked = Signal(int, str)
    page_double_clicked = Signal(int, str)
    file_dropped = Signal(str)
    exclusion_zone_drawn = Signal(str, int, float, float, float, float)  # side, page_idx, x, y, w, h

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
            if thumb._original_pixmap is not None or thumb._highlight_pixmap is not None:
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
            thumb.exclusion_zone_drawn.connect(self._on_exclusion_zone_drawn)

            if page.thumbnail:
                pixmap = pil_to_qpixmap(page.thumbnail)
                thumb.set_pixmap(pixmap)

            self.thumbnails.append(thumb)
            self.layout.insertWidget(self.layout.count() - 1, thumb)

    def _on_exclusion_zone_drawn(self, side: str, page_index: int, x: float, y: float, w: float, h: float) -> None:
        """Forward exclusion zone signal to parent."""
        self.exclusion_zone_drawn.emit(side, page_index, x, y, w, h)

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

    def set_diff_result(self, page_index: int, diff_result: DiffResult, highlight_pixmap: Optional[QPixmap] = None) -> None:
        """Set diff result for a specific page thumbnail."""
        if page_index < len(self.thumbnails):
            self.thumbnails[page_index].set_diff_result(diff_result, highlight_pixmap)

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

    def scroll_to_page(self, page_index: int) -> None:
        """Scroll to make a specific page visible."""
        if page_index < 0 or page_index >= len(self.thumbnails):
            return

        thumb = self.thumbnails[page_index]
        # Ensure the widget is scrolled into view, centered if possible
        self.ensureWidgetVisible(thumb, 0, self.height() // 4)

    def set_drawing_mode(self, enabled: bool) -> None:
        """Enable or disable exclusion zone drawing mode on all thumbnails."""
        for thumb in self.thumbnails:
            thumb.set_drawing_mode(enabled)

    def update_exclusion_zones(self, zones: List[ExclusionZone]) -> None:
        """Update exclusion zone overlays on all thumbnails."""
        for thumb in self.thumbnails:
            applicable_zones = []
            for z in zones:
                if z.enabled:
                    if z.applies_to == AppliesTo.BOTH:
                        applicable_zones.append((z.x, z.y, z.width, z.height))
                    elif z.applies_to == AppliesTo.LEFT and self.side == "left":
                        applicable_zones.append((z.x, z.y, z.width, z.height))
                    elif z.applies_to == AppliesTo.RIGHT and self.side == "right":
                        applicable_zones.append((z.x, z.y, z.width, z.height))
            thumb.set_exclusion_zones(applicable_zones)

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


class ExclusionZoneDialog(QDialog):
    """Dialog for managing exclusion zones."""

    def __init__(self, parent: QWidget, zone_set: ExclusionZoneSet):
        super().__init__(parent)
        self.zone_set = zone_set
        self.setWindowTitle("除外領域")
        self.setMinimumSize(500, 400)
        self._setup_ui()
        self._populate_list()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        layout = QVBoxLayout(self)

        # Zone list
        layout.addWidget(QLabel("除外領域一覧:"))
        self.zone_list = QListWidget()
        self.zone_list.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.zone_list)

        # Preset buttons
        preset_group = QGroupBox("プリセット追加")
        preset_layout = QHBoxLayout(preset_group)

        presets = [
            ("ページ番号(下)", ExclusionZoneSet.preset_page_number_bottom),
            ("ページ番号(右下)", ExclusionZoneSet.preset_page_number_bottom_right),
            ("ヘッダー", ExclusionZoneSet.preset_header),
            ("フッター", ExclusionZoneSet.preset_footer),
            ("スライド番号", ExclusionZoneSet.preset_slide_number_ppt),
        ]

        for name, factory in presets:
            btn = QPushButton(name)
            btn.clicked.connect(lambda checked, f=factory: self._add_preset(f))
            preset_layout.addWidget(btn)

        layout.addWidget(preset_group)

        # Manual input
        manual_group = QGroupBox("カスタム領域追加")
        manual_layout = QFormLayout(manual_group)

        coord_layout = QHBoxLayout()
        self.x_input = QSpinBox()
        self.x_input.setRange(0, 100)
        coord_layout.addWidget(QLabel("X:"))
        coord_layout.addWidget(self.x_input)

        self.y_input = QSpinBox()
        self.y_input.setRange(0, 100)
        coord_layout.addWidget(QLabel("Y:"))
        coord_layout.addWidget(self.y_input)

        self.w_input = QSpinBox()
        self.w_input.setRange(1, 100)
        self.w_input.setValue(20)
        coord_layout.addWidget(QLabel("幅:"))
        coord_layout.addWidget(self.w_input)

        self.h_input = QSpinBox()
        self.h_input.setRange(1, 100)
        self.h_input.setValue(10)
        coord_layout.addWidget(QLabel("高さ:"))
        coord_layout.addWidget(self.h_input)

        manual_layout.addRow("座標 (%):", coord_layout)

        name_layout = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("領域名")
        name_layout.addWidget(self.name_input)

        self.applies_combo = QComboBox()
        self.applies_combo.addItems(["両方", "左のみ", "右のみ"])
        name_layout.addWidget(self.applies_combo)

        add_btn = QPushButton("追加")
        add_btn.clicked.connect(self._add_custom_zone)
        name_layout.addWidget(add_btn)

        manual_layout.addRow("名前/適用:", name_layout)
        layout.addWidget(manual_group)

        # Remove button
        remove_btn = QPushButton("選択した領域を削除")
        remove_btn.clicked.connect(self._remove_zone)
        layout.addWidget(remove_btn)

        # Dialog buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _populate_list(self) -> None:
        """Populate the list with current zones."""
        self.zone_list.clear()
        for zone in self.zone_set.zones:
            label = f"{zone.name or '名前なし'} ({zone.x:.0%}, {zone.y:.0%}, {zone.width:.0%}x{zone.height:.0%})"
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if zone.enabled else Qt.CheckState.Unchecked)
            self.zone_list.addItem(item)

    def _add_preset(self, factory) -> None:
        """Add a preset zone."""
        zone = factory()
        self.zone_set.add(zone)
        self._populate_list()

    def _add_custom_zone(self) -> None:
        """Add a custom zone from inputs."""
        x = self.x_input.value() / 100.0
        y = self.y_input.value() / 100.0
        w = self.w_input.value() / 100.0
        h = self.h_input.value() / 100.0
        name = self.name_input.text() or "カスタム"

        applies_map = {0: AppliesTo.BOTH, 1: AppliesTo.LEFT, 2: AppliesTo.RIGHT}
        applies_to = applies_map[self.applies_combo.currentIndex()]

        zone = ExclusionZone(x=x, y=y, width=w, height=h, name=name, applies_to=applies_to)
        self.zone_set.add(zone)
        self._populate_list()

    def _remove_zone(self) -> None:
        """Remove the selected zone."""
        row = self.zone_list.currentRow()
        if row >= 0 and row < len(self.zone_set.zones):
            zone = self.zone_set.zones[row]
            self.zone_set.remove(zone)
            self._populate_list()

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        """Handle checkbox toggle."""
        row = self.zone_list.row(item)
        if row >= 0 and row < len(self.zone_set.zones):
            self.zone_set.zones[row].enabled = item.checkState() == Qt.CheckState.Checked


class DiffSummaryDialog(QWidget):
    """Modeless dialog showing pages with differences."""

    jump_requested = Signal(int, int)  # left_idx, right_idx

    def __init__(self, parent: QWidget):
        super().__init__(parent, Qt.WindowType.Window)
        self.setWindowTitle("差分一覧")
        self.setMinimumSize(400, 500)
        self.diff_pages: List[tuple] = []
        self._setup_ui()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("差分があるページ一覧")
        font = header.font()
        font.setPointSize(12)
        font.setBold(True)
        header.setFont(font)
        layout.addWidget(header)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["左ページ", "右ページ", "状態"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.cellDoubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table)

        # Summary
        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        # Buttons
        btn_layout = QHBoxLayout()
        jump_btn = QPushButton("ジャンプ")
        jump_btn.clicked.connect(self._on_jump)
        btn_layout.addWidget(jump_btn)

        export_btn = QPushButton("CSVエクスポート")
        export_btn.clicked.connect(self._on_export)
        btn_layout.addWidget(export_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def update_diff_list(self, matching_result: Optional[MatchingResult], diff_scores: dict) -> None:
        """Update the list with current diff information."""
        self.table.setRowCount(0)
        self.diff_pages = []

        if not matching_result:
            self.summary_label.setText("比較結果がありません")
            return

        diff_count = 0
        identical_count = 0
        unmatched_count = 0

        for match in matching_result.matches:
            if match.status == MatchStatus.MATCHED:
                has_diff = diff_scores.get((match.left_index, match.right_index), False)
                if has_diff:
                    diff_count += 1
                    status = "差分あり"
                    color = QColor(255, 200, 200)
                else:
                    identical_count += 1
                    status = "同一"
                    color = QColor(200, 255, 200)

                self.diff_pages.append((match.left_index, match.right_index, has_diff))

                row = self.table.rowCount()
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(str(match.left_index + 1)))
                self.table.setItem(row, 1, QTableWidgetItem(str(match.right_index + 1)))
                self.table.setItem(row, 2, QTableWidgetItem(status))

                for col in range(3):
                    self.table.item(row, col).setBackground(color)
            else:
                unmatched_count += 1

        self.summary_label.setText(
            f"差分あり: {diff_count}  同一: {identical_count}  未マッチ: {unmatched_count}"
        )

    def _on_double_click(self, row: int, col: int) -> None:
        """Handle double click on row."""
        self._jump_to_row(row)

    def _on_jump(self) -> None:
        """Jump to selected row."""
        row = self.table.currentRow()
        if row >= 0:
            self._jump_to_row(row)

    def _jump_to_row(self, row: int) -> None:
        """Jump to the specified row."""
        if row >= 0 and row < len(self.diff_pages):
            left_idx, right_idx, _ = self.diff_pages[row]
            self.jump_requested.emit(left_idx, right_idx)

    def _on_export(self) -> None:
        """Export diff list to CSV."""
        path, _ = QFileDialog.getSaveFileName(
            self, "CSVエクスポート", "", "CSV Files (*.csv)"
        )
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write("Left Page,Right Page,Status\n")
                    for left_idx, right_idx, has_diff in self.diff_pages:
                        status = "差分あり" if has_diff else "同一"
                        f.write(f"{left_idx + 1},{right_idx + 1},{status}\n")
                QMessageBox.information(self, "完了", f"エクスポート完了: {path}")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"エクスポート失敗: {e}")

    def get_diff_page_indices(self) -> List[tuple]:
        """Get list of (left_idx, right_idx) for pages with differences."""
        return [(left, right) for left, right, has_diff in self.diff_pages if has_diff]


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

        for left_pos, right_pos, status, has_diff in self.links:
            if left_pos is None or right_pos is None:
                continue

            # Choose color based on has_diff (True = has differences, False = identical)
            if status == MatchStatus.MATCHED:
                if has_diff:
                    color = QColor(220, 53, 69)  # Red - has differences
                else:
                    color = QColor(40, 167, 69)  # Green - no difference
            else:
                color = QColor(108, 117, 125)  # Gray - unmatched

            pen = QPen(color)
            pen.setWidth(8)  # Thick line for visibility
            painter.setPen(pen)

            # Draw line
            painter.drawLine(left_pos, right_pos)

        painter.end()


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT/PDF Comparator (PySide6)")
        self.setMinimumSize(800, 600)
        self.resize(1600, 1000)

        self.session = Session()
        self.left_doc: Optional[Document] = None
        self.right_doc: Optional[Document] = None
        self.matching_result: Optional[MatchingResult] = None
        self.exclusion_zones = ExclusionZoneSet()
        self.diff_scores: dict = {}  # (left_idx, right_idx) -> has_differences

        self._selected_left: Optional[int] = None
        self._selected_right: Optional[int] = None
        self._current_diff_index: int = -1  # Current position in diff navigation
        self._highlight_color: tuple = (255, 0, 0)  # RGB for highlight
        self._drawing_mode: bool = False  # Exclusion zone drawing mode

        # Modeless dialogs
        self._diff_summary_dialog: Optional[DiffSummaryDialog] = None

        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()

        # Enable drag and drop on main window
        self.setAcceptDrops(True)

        # Show maximized after all initialization
        self.showMaximized()

    def _setup_ui(self) -> None:
        """Set up the main UI layout."""
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Content area
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

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

        # Connect exclusion zone signals
        self.left_panel.exclusion_zone_drawn.connect(self._on_exclusion_zone_drawn)
        self.right_panel.exclusion_zone_drawn.connect(self._on_exclusion_zone_drawn)

        self.splitter.addWidget(self.left_panel)
        self.splitter.addWidget(self.right_panel)
        self.splitter.setSizes([800, 800])

        content_layout.addWidget(self.splitter)
        main_layout.addWidget(content_widget, 1)

        # Bottom sync scroll bar area
        sync_widget = QWidget()
        sync_widget.setStyleSheet("background-color: #dcdcdc;")
        sync_widget.setMinimumHeight(40)
        sync_layout = QHBoxLayout(sync_widget)
        sync_layout.setContentsMargins(10, 5, 10, 5)

        sync_label = QLabel("連動スクロール:")
        sync_label.setStyleSheet("color: #505050;")
        sync_layout.addWidget(sync_label)

        self.sync_scrollbar = QSlider(Qt.Orientation.Horizontal)
        self.sync_scrollbar.setMinimum(0)
        self.sync_scrollbar.setMaximum(100)
        self.sync_scrollbar.setValue(0)
        self.sync_scrollbar.valueChanged.connect(self._on_sync_scroll)
        sync_layout.addWidget(self.sync_scrollbar, 1)

        self.sync_checkbox = QCheckBox("連動")
        self.sync_checkbox.setChecked(False)
        self.sync_checkbox.stateChanged.connect(self._on_sync_toggle)
        sync_layout.addWidget(self.sync_checkbox)

        main_layout.addWidget(sync_widget)

        # Link overlay (positioned over entire window)
        self.link_overlay = LinkOverlay(content_widget)

        # Update overlay on scroll
        self.left_panel.verticalScrollBar().valueChanged.connect(self._on_panel_scroll)
        self.right_panel.verticalScrollBar().valueChanged.connect(self._on_panel_scroll)

        # Timer for delayed link updates
        self.link_update_timer = QTimer()
        self.link_update_timer.setSingleShot(True)
        self.link_update_timer.timeout.connect(self._do_update_links)

    def _on_panel_scroll(self) -> None:
        """Handle scroll event from document panels."""
        # Sync scroll if enabled
        if self.sync_checkbox.isChecked():
            sender = self.sender()
            if sender == self.left_panel.verticalScrollBar():
                self._sync_scroll_from(self.left_panel, self.right_panel)
            elif sender == self.right_panel.verticalScrollBar():
                self._sync_scroll_from(self.right_panel, self.left_panel)

        self._update_links()
        self._update_sync_scrollbar()

    def _on_sync_scroll(self, value: int) -> None:
        """Handle sync scrollbar movement - scroll through matched pairs."""
        if not self.left_doc or not self.right_doc:
            return

        # If we have matching results, scroll through matched pairs
        if self.matching_result:
            matched_pairs = self.matching_result.get_matched_pairs()
            if matched_pairs:
                # Map slider value to pair index
                pair_idx = int(value / 100.0 * len(matched_pairs))
                pair_idx = min(pair_idx, len(matched_pairs) - 1)

                left_idx, right_idx, _ = matched_pairs[pair_idx]
                self.left_panel.scroll_to_page(left_idx)
                self.right_panel.scroll_to_page(right_idx)

                self._update_links()
                return

        # Fallback: scroll both panels to the same percentage
        percent = value / 100.0
        self._scroll_panel_to_percent(self.left_panel, percent)
        self._scroll_panel_to_percent(self.right_panel, percent)
        self._update_links()

    def _on_sync_toggle(self, state: int) -> None:
        """Handle sync checkbox toggle."""
        if state:
            self._update_sync_scrollbar()

    def _sync_scroll_from(self, source: DocumentPanel, target: DocumentPanel) -> None:
        """Sync scroll position from source to target panel, trying to align paired slides."""
        if not self.matching_result:
            # Fallback to percentage-based sync if no matching
            source_bar = source.verticalScrollBar()
            target_bar = target.verticalScrollBar()
            source_range = source_bar.maximum()
            if source_range > 0:
                percent = source_bar.value() / source_range
                target_pos = int(percent * target_bar.maximum())
                target_bar.setValue(target_pos)
            return

        # Find the most visible slide in source panel
        source_visible_idx = self._get_center_visible_page(source)
        if source_visible_idx is None:
            return

        # Find paired slide
        if source == self.left_panel:
            match = self.matching_result.get_match_for_left(source_visible_idx)
            if match and match.right_index is not None:
                target.scroll_to_page(match.right_index)
        else:
            match = self.matching_result.get_match_for_right(source_visible_idx)
            if match and match.left_index is not None:
                target.scroll_to_page(match.left_index)

    def _get_center_visible_page(self, panel: DocumentPanel) -> Optional[int]:
        """Get the page index that's most visible (closest to center) in the panel."""
        if not panel.thumbnails:
            return None

        viewport = panel.viewport()
        if not viewport:
            return None
        panel_height = viewport.height()
        panel_center_y = panel_height // 2

        # Find which thumbnail is closest to the center
        best_idx = 0
        best_distance = float('inf')

        for i, thumb in enumerate(panel.thumbnails):
            # Get thumbnail position relative to viewport
            thumb_pos = thumb.mapTo(viewport, QPoint(0, 0))
            thumb_center_y = thumb_pos.y() + thumb.height() // 2

            distance = abs(thumb_center_y - panel_center_y)
            if distance < best_distance:
                best_distance = distance
                best_idx = i

        return best_idx

    def _scroll_panel_to_percent(self, panel: DocumentPanel, percent: float) -> None:
        """Scroll panel to a percentage position."""
        scrollbar = panel.verticalScrollBar()
        pos = int(percent * scrollbar.maximum())
        scrollbar.setValue(pos)

    def _update_sync_scrollbar(self) -> None:
        """Update sync scrollbar position based on currently visible pair."""
        # If we have matching results, position based on current pair
        if self.matching_result:
            matched_pairs = self.matching_result.get_matched_pairs()
            if matched_pairs:
                visible_idx = self._get_center_visible_page(self.left_panel)
                if visible_idx is not None:
                    # Find which pair contains this index
                    for i, (left_idx, right_idx, _) in enumerate(matched_pairs):
                        if left_idx == visible_idx:
                            percent = int((i / len(matched_pairs)) * 100)
                            self.sync_scrollbar.blockSignals(True)
                            self.sync_scrollbar.setValue(min(100, max(0, percent)))
                            self.sync_scrollbar.blockSignals(False)
                            return

        # Fallback: use scroll percentage
        scrollbar = self.left_panel.verticalScrollBar()
        if scrollbar.maximum() > 0:
            percent = int((scrollbar.value() / scrollbar.maximum()) * 100)
            self.sync_scrollbar.blockSignals(True)
            self.sync_scrollbar.setValue(min(100, max(0, percent)))
            self.sync_scrollbar.blockSignals(False)

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

        exclusion_zones = QAction("除外領域...", self)
        exclusion_zones.triggered.connect(self._show_exclusion_zones_dialog)
        compare_menu.addAction(exclusion_zones)

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

        # Diff navigation buttons
        prev_diff_btn = QAction("← Prev Diff", self)
        prev_diff_btn.setShortcut("Ctrl+Up")
        prev_diff_btn.triggered.connect(self._go_prev_diff)
        toolbar.addAction(prev_diff_btn)

        next_diff_btn = QAction("Next Diff →", self)
        next_diff_btn.setShortcut("Ctrl+Down")
        next_diff_btn.triggered.connect(self._go_next_diff)
        toolbar.addAction(next_diff_btn)

        toolbar.addSeparator()

        # Diff summary button
        diff_summary_btn = QAction("差分一覧", self)
        diff_summary_btn.triggered.connect(self._show_diff_summary)
        toolbar.addAction(diff_summary_btn)

        toolbar.addSeparator()

        # Color picker button
        color_btn = QAction("ハイライト色", self)
        color_btn.triggered.connect(self._pick_highlight_color)
        toolbar.addAction(color_btn)

        toolbar.addSeparator()

        # Drawing mode toggle
        self._draw_mode_action = QAction("除外領域描画", self)
        self._draw_mode_action.setCheckable(True)
        self._draw_mode_action.triggered.connect(self._toggle_drawing_mode)
        toolbar.addAction(self._draw_mode_action)

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

            # Auto-compare if both documents are loaded
            if self.left_doc and self.right_doc:
                QTimer.singleShot(100, self._run_comparison)

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

        # Phase 1: Match pages
        matcher = PageMatcher()
        self.matching_result = matcher.match(
            self.left_doc, self.right_doc,
            progress_callback=progress_callback
        )

        # Phase 2: Compute differences for matched pairs
        progress.setLabelText("Computing differences...")
        comparator = ImageComparator(highlight_color=self._highlight_color)
        matched_pairs = self.matching_result.get_matched_pairs()
        total_pairs = len(matched_pairs)
        self.diff_scores = {}  # Clear previous scores

        for i, (left_idx, right_idx, score) in enumerate(matched_pairs):
            progress.setValue(int((i + 1) / total_pairs * 100) if total_pairs > 0 else 100)
            progress.setLabelText(f"Computing diff {i + 1}/{total_pairs}...")
            QApplication.processEvents()

            left_page = self.left_doc.pages[left_idx]
            right_page = self.right_doc.pages[right_idx]

            if left_page.thumbnail and right_page.thumbnail:
                # Get all exclusion zones (applied to both sides)
                all_zones = self.exclusion_zones.zones

                # Compare the images with all exclusion zones
                diff_result = comparator.compare(
                    left_page.thumbnail, right_page.thumbnail,
                    exclusion_zones=all_zones
                )

                # Store whether there are differences (for link coloring)
                self.diff_scores[(left_idx, right_idx)] = diff_result.has_differences

                # Apply highlighted images to thumbnails
                if diff_result.highlight_image:
                    left_highlight_pixmap = pil_to_qpixmap(diff_result.highlight_image)
                    self.left_panel.set_diff_result(left_idx, diff_result, left_highlight_pixmap)

                    # Also create highlight for right side
                    diff_result_right = comparator.compare(
                        right_page.thumbnail, left_page.thumbnail,
                        exclusion_zones=all_zones
                    )
                    if diff_result_right.highlight_image:
                        right_highlight_pixmap = pil_to_qpixmap(diff_result_right.highlight_image)
                        self.right_panel.set_diff_result(right_idx, diff_result_right, right_highlight_pixmap)

        progress.close()

        self.session.matching_result = self.matching_result
        self._current_diff_index = -1  # Reset diff navigation

        # Update UI
        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_links()
        self._update_status()

        # Update diff summary dialog if open
        if self._diff_summary_dialog and self._diff_summary_dialog.isVisible():
            self._diff_summary_dialog.update_diff_list(self.matching_result, self.diff_scores)

    def _on_page_clicked(self, index: int, side: str) -> None:
        """Handle page click for manual linking and scroll to paired slide."""
        if side == "left":
            self._selected_left = index
            self.left_panel.set_selected(index)
            # Scroll to paired slide on the right
            self._scroll_to_paired_slide(index, "left")
        else:
            self._selected_right = index
            self.right_panel.set_selected(index)
            # Scroll to paired slide on the left
            self._scroll_to_paired_slide(index, "right")

        # If both selected, create manual link
        if self._selected_left is not None and self._selected_right is not None:
            self._create_manual_link(self._selected_left, self._selected_right)
            self._selected_left = None
            self._selected_right = None
            self.left_panel.set_selected(None)
            self.right_panel.set_selected(None)

    def _scroll_to_paired_slide(self, clicked_index: int, clicked_side: str) -> None:
        """Scroll the opposite panel to show the paired slide."""
        if not self.matching_result:
            return

        # Find the matched page
        if clicked_side == "left":
            match = self.matching_result.get_match_for_left(clicked_index)
            if match and match.right_index is not None:
                self.right_panel.scroll_to_page(match.right_index)
        else:
            match = self.matching_result.get_match_for_right(clicked_index)
            if match and match.left_index is not None:
                self.left_panel.scroll_to_page(match.left_index)

    def _on_page_double_clicked(self, index: int, side: str) -> None:
        """Handle page double click for viewing details."""
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
                    # Use has_differences for coloring (True = red, False = green)
                    has_diff = self.diff_scores.get((match.left_index, match.right_index), False)
                    links.append((left_pos, right_pos, match.status, has_diff))

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

    def _show_exclusion_zones_dialog(self) -> None:
        """Show the exclusion zones dialog."""
        dialog = ExclusionZoneDialog(self, self.exclusion_zones)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.session.exclusion_zones = self.exclusion_zones
            self._update_exclusion_zone_overlays()
            self.statusbar.showMessage(
                f"除外領域を更新しました ({len(self.exclusion_zones.zones)} 領域)"
            )

    def _toggle_drawing_mode(self, checked: bool) -> None:
        """Toggle exclusion zone drawing mode."""
        self._drawing_mode = checked
        self.left_panel.set_drawing_mode(checked)
        self.right_panel.set_drawing_mode(checked)

        if checked:
            self.statusbar.showMessage("除外領域描画モード: 画像上をドラッグして除外領域を指定してください")
        else:
            self.statusbar.showMessage("除外領域描画モードを終了しました")
            self._update_exclusion_zone_overlays()

    def _on_exclusion_zone_drawn(self, side: str, page_index: int, x: float, y: float, w: float, h: float) -> None:
        """Handle exclusion zone drawn on a thumbnail."""
        # Always apply to BOTH sides for consistent comparison
        zone_count = len(self.exclusion_zones.zones) + 1
        zone = ExclusionZone(
            x=x, y=y, width=w, height=h,
            name=f"手動領域 {zone_count}",
            applies_to=AppliesTo.BOTH
        )
        self.exclusion_zones.add(zone)
        self.session.exclusion_zones = self.exclusion_zones
        self._update_exclusion_zone_overlays()

        self.statusbar.showMessage(
            f"除外領域を追加しました: {zone.name} ({x*100:.0f}%, {y*100:.0f}%, {w*100:.0f}%x{h*100:.0f}%)"
        )

    def _update_exclusion_zone_overlays(self) -> None:
        """Update exclusion zone overlays on both panels."""
        left_zones = self.exclusion_zones.get_zones_for("left")
        right_zones = self.exclusion_zones.get_zones_for("right")
        self.left_panel.update_exclusion_zones(left_zones)
        self.right_panel.update_exclusion_zones(right_zones)

    def _get_diff_pages(self) -> List[tuple]:
        """Get list of (left_idx, right_idx) for pages with differences."""
        if not self.matching_result:
            return []
        diff_pages = []
        for match in self.matching_result.matches:
            if match.status == MatchStatus.MATCHED:
                has_diff = self.diff_scores.get((match.left_index, match.right_index), False)
                if has_diff:
                    diff_pages.append((match.left_index, match.right_index))
        return diff_pages

    def _go_prev_diff(self) -> None:
        """Navigate to previous page with differences."""
        diff_pages = self._get_diff_pages()
        if not diff_pages:
            self.statusbar.showMessage("差分がありません")
            return

        if self._current_diff_index <= 0:
            self._current_diff_index = len(diff_pages) - 1
        else:
            self._current_diff_index -= 1

        left_idx, right_idx = diff_pages[self._current_diff_index]
        self.jump_to_page_pair(left_idx, right_idx)
        self.statusbar.showMessage(f"差分 {self._current_diff_index + 1}/{len(diff_pages)}")

    def _go_next_diff(self) -> None:
        """Navigate to next page with differences."""
        diff_pages = self._get_diff_pages()
        if not diff_pages:
            self.statusbar.showMessage("差分がありません")
            return

        if self._current_diff_index >= len(diff_pages) - 1:
            self._current_diff_index = 0
        else:
            self._current_diff_index += 1

        left_idx, right_idx = diff_pages[self._current_diff_index]
        self.jump_to_page_pair(left_idx, right_idx)
        self.statusbar.showMessage(f"差分 {self._current_diff_index + 1}/{len(diff_pages)}")

    def jump_to_page_pair(self, left_idx: int, right_idx: int) -> None:
        """Jump to a specific page pair."""
        self.left_panel.scroll_to_page(left_idx)
        self.right_panel.scroll_to_page(right_idx)
        self.left_panel.set_selected(left_idx)
        self.right_panel.set_selected(right_idx)
        QTimer.singleShot(50, self._do_update_links)

    def _show_diff_summary(self) -> None:
        """Show the diff summary dialog."""
        if self._diff_summary_dialog is None:
            self._diff_summary_dialog = DiffSummaryDialog(self)
            self._diff_summary_dialog.jump_requested.connect(self.jump_to_page_pair)

        self._diff_summary_dialog.update_diff_list(self.matching_result, self.diff_scores)
        self._diff_summary_dialog.show()
        self._diff_summary_dialog.raise_()
        self._diff_summary_dialog.activateWindow()

    def _pick_highlight_color(self) -> None:
        """Open color picker for highlight color."""
        current = QColor(*self._highlight_color)
        color = QColorDialog.getColor(current, self, "ハイライト色を選択")

        if color.isValid():
            self._highlight_color = (color.red(), color.green(), color.blue())
            self.statusbar.showMessage(f"ハイライト色を変更しました: RGB{self._highlight_color}")

            # Re-run comparison with new color if documents are loaded
            if self.left_doc and self.right_doc and self.matching_result:
                self._run_comparison()

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

    def closeEvent(self, event) -> None:
        """Handle window close event."""
        # Close diff summary dialog if open
        if self._diff_summary_dialog:
            self._diff_summary_dialog.close()
            self._diff_summary_dialog = None

        # Close PowerPoint cache if used
        try:
            from ..core.document import close_powerpoint_cache
            close_powerpoint_cache()
        except Exception:
            pass

        super().closeEvent(event)
