"""Main window for wxPython GUI."""

from __future__ import annotations

from pathlib import Path
from typing import Optional, List
import wx
import wx.lib.scrolledpanel as scrolled

try:
    from src.core.document import Document
    from src.core.page_matcher import PageMatcher, MatchingResult, MatchStatus
    from src.core.image_comparator import ImageComparator, DiffResult
    from src.core.exclusion_zone import ExclusionZone, ExclusionZoneSet, AppliesTo
    from src.core.session import Session
    from src.utils.image_utils import pil_to_wxbitmap
except ImportError:
    from core.document import Document
    from core.page_matcher import PageMatcher, MatchingResult, MatchStatus
    from core.image_comparator import ImageComparator, DiffResult
    from core.exclusion_zone import ExclusionZone, ExclusionZoneSet, AppliesTo
    from core.session import Session
    from utils.image_utils import pil_to_wxbitmap


class PageThumbnailPanel(wx.Panel):
    """Panel displaying a single page thumbnail."""

    def __init__(
        self,
        parent: wx.Window,
        page_index: int,
        side: str,
    ):
        super().__init__(parent, style=wx.BORDER_SIMPLE)
        self.page_index = page_index
        self.side = side
        self._original_bitmap: Optional[wx.Bitmap] = None
        self._highlight_bitmap: Optional[wx.Bitmap] = None  # Diff highlighted version
        self._show_diff: bool = True  # Whether to show diff highlights
        self._selected = False
        self._match_status: Optional[MatchStatus] = None
        self._diff_result: Optional[DiffResult] = None

        # Exclusion zone drawing state
        self._drawing_mode: bool = False
        self._draw_start: Optional[wx.Point] = None
        self._draw_current: Optional[wx.Point] = None
        self._exclusion_zones: List[tuple] = []  # List of (x, y, w, h) in normalized coords
        self._selected_zone_index: int = -1  # Index of selected zone

        self.SetBackgroundColour(wx.WHITE)

        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Image display - use a panel for custom drawing
        self.image_panel = wx.Panel(self)
        self.image_panel.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.sizer.Add(self.image_panel, 1, wx.ALL | wx.EXPAND, 5)

        # Page label
        self.page_label = wx.StaticText(self, label=f"Page {page_index + 1}")
        self.page_label.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        self.sizer.Add(self.page_label, 0, wx.ALL | wx.CENTER, 2)

        self.SetSizer(self.sizer)

        # For backward compatibility, keep image_ctrl as alias
        self.image_ctrl = self.image_panel
        self._scaled_bitmap: Optional[wx.Bitmap] = None

        # Bind events
        self.Bind(wx.EVT_LEFT_DOWN, self._on_click)
        self.Bind(wx.EVT_LEFT_DCLICK, self._on_double_click)
        self.Bind(wx.EVT_SIZE, self._on_resize)
        self.image_panel.Bind(wx.EVT_LEFT_DOWN, self._on_image_click)
        self.image_panel.Bind(wx.EVT_LEFT_UP, self._on_image_release)
        self.image_panel.Bind(wx.EVT_MOTION, self._on_image_motion)
        self.image_panel.Bind(wx.EVT_PAINT, self._on_image_paint)
        self.image_panel.Bind(wx.EVT_LEFT_DCLICK, self._on_double_click)
        self.image_panel.Bind(wx.EVT_RIGHT_DOWN, self._on_image_right_click)

    def set_bitmap(self, bitmap: wx.Bitmap) -> None:
        """Set the thumbnail image."""
        self._original_bitmap = bitmap
        self._highlight_bitmap = None  # Clear highlight when original changes
        self._update_scaled_bitmap()

    def set_diff_result(self, diff_result: Optional[DiffResult], highlight_image: Optional[wx.Bitmap] = None) -> None:
        """Set the diff result and highlighted image."""
        self._diff_result = diff_result
        self._highlight_bitmap = highlight_image

        # Update scaled bitmap and refresh
        if self._show_diff and self._highlight_bitmap is not None:
            img = self._highlight_bitmap.ConvertToImage()
            parent = self.GetParent()
            if parent:
                parent_width = parent.GetClientSize().width
                target_w = max(100, parent_width - 50)
                orig_w, orig_h = img.GetWidth(), img.GetHeight()
                scale = target_w / orig_w
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)
                img = img.Scale(new_w, new_h, wx.IMAGE_QUALITY_HIGH)
                self._scaled_bitmap = wx.Bitmap(img)

        self.Layout()
        self.image_panel.Refresh()
        self.Update()

    def set_drawing_mode(self, enabled: bool) -> None:
        """Enable or disable exclusion zone drawing mode."""
        self._drawing_mode = enabled
        if enabled:
            self.image_panel.SetCursor(wx.Cursor(wx.CURSOR_CROSS))
        else:
            self.image_panel.SetCursor(wx.NullCursor)
        self._draw_start = None
        self._draw_current = None
        self.image_panel.Refresh()

    def set_exclusion_zones(self, zones: List[tuple]) -> None:
        """Set exclusion zones to display as overlays."""
        self._exclusion_zones = zones
        self.image_panel.Refresh()

    def _on_image_paint(self, event: wx.PaintEvent) -> None:
        """Paint the image with exclusion zone overlays."""
        dc = wx.AutoBufferedPaintDC(self.image_panel)
        dc.Clear()

        # Draw the scaled bitmap
        if self._scaled_bitmap:
            dc.DrawBitmap(self._scaled_bitmap, 0, 0)

        # Get image size for coordinate conversion
        if not self._scaled_bitmap:
            return
        img_w = self._scaled_bitmap.GetWidth()
        img_h = self._scaled_bitmap.GetHeight()

        # Draw existing exclusion zones
        for i, (x, y, w, h) in enumerate(self._exclusion_zones):
            rect_x = int(x * img_w)
            rect_y = int(y * img_h)
            rect_w = int(w * img_w)
            rect_h = int(h * img_h)
            if i == self._selected_zone_index:
                # Selected zone - yellow highlight
                dc.SetBrush(wx.Brush(wx.Colour(255, 255, 0, 120), wx.BRUSHSTYLE_SOLID))
                dc.SetPen(wx.Pen(wx.Colour(255, 200, 0), 3, wx.PENSTYLE_SOLID))
            else:
                # Normal zone - red
                dc.SetBrush(wx.Brush(wx.Colour(255, 0, 0, 64), wx.BRUSHSTYLE_BDIAGONAL_HATCH))
                dc.SetPen(wx.Pen(wx.Colour(255, 0, 0), 2, wx.PENSTYLE_DOT))
            dc.DrawRectangle(rect_x, rect_y, rect_w, rect_h)

        # Draw current selection rectangle
        if self._drawing_mode and self._draw_start and self._draw_current:
            dc.SetBrush(wx.Brush(wx.Colour(0, 120, 215, 50)))
            dc.SetPen(wx.Pen(wx.Colour(0, 120, 215), 2))
            x1, y1 = self._draw_start
            x2, y2 = self._draw_current
            rect = wx.Rect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
            dc.DrawRectangle(rect)

    def _get_zone_at_pos(self, pos: wx.Point) -> int:
        """Return index of zone at position, or -1 if none."""
        if not self._scaled_bitmap:
            return -1
        img_w = self._scaled_bitmap.GetWidth()
        img_h = self._scaled_bitmap.GetHeight()
        px, py = pos.x, pos.y
        for i, (x, y, w, h) in enumerate(self._exclusion_zones):
            rx = int(x * img_w)
            ry = int(y * img_h)
            rw = int(w * img_w)
            rh = int(h * img_h)
            if rx <= px <= rx + rw and ry <= py <= ry + rh:
                return i
        return -1

    def _on_image_click(self, event: wx.MouseEvent) -> None:
        """Handle click on image panel."""
        if self._drawing_mode:
            self._draw_start = event.GetPosition()
            self._draw_current = self._draw_start
            self.image_panel.CaptureMouse()
        else:
            # Check if clicking on an exclusion zone
            pos = event.GetPosition()
            zone_idx = self._get_zone_at_pos(pos)
            if zone_idx >= 0:
                self._selected_zone_index = zone_idx
                self.image_panel.Refresh()
            else:
                self._selected_zone_index = -1
                self.image_panel.Refresh()
                self._on_click(event)

    def _on_image_right_click(self, event: wx.MouseEvent) -> None:
        """Handle right click on image panel - context menu for zones."""
        pos = event.GetPosition()
        zone_idx = self._get_zone_at_pos(pos)
        if zone_idx >= 0:
            self._selected_zone_index = zone_idx
            self.image_panel.Refresh()
            # Show context menu
            menu = wx.Menu()
            delete_item = menu.Append(wx.ID_ANY, "この除外領域を削除")
            self.Bind(wx.EVT_MENU, lambda e: self._notify_zone_delete(zone_idx), delete_item)
            self.PopupMenu(menu)
            menu.Destroy()

    def _notify_zone_delete(self, zone_idx: int) -> None:
        """Notify parent to delete a zone."""
        parent = self.GetParent()
        while parent:
            if hasattr(parent, 'on_exclusion_zone_delete'):
                parent.on_exclusion_zone_delete(zone_idx)
                return
            parent = parent.GetParent()

    def _on_image_release(self, event: wx.MouseEvent) -> None:
        """Handle mouse release on image panel."""
        if self.image_panel.HasCapture():
            self.image_panel.ReleaseMouse()

        if self._drawing_mode and self._draw_start and self._draw_current:
            # Calculate normalized coordinates
            if self._scaled_bitmap:
                img_w = self._scaled_bitmap.GetWidth()
                img_h = self._scaled_bitmap.GetHeight()

                x1, y1 = self._draw_start
                x2, y2 = self._draw_current

                # Normalize to 0-1 range
                nx = min(x1, x2) / img_w
                ny = min(y1, y2) / img_h
                nw = abs(x2 - x1) / img_w
                nh = abs(y2 - y1) / img_h

                # Only add if zone is large enough (at least 1% in each dimension)
                if nw > 0.01 and nh > 0.01:
                    # Notify parent to add exclusion zone
                    self._notify_exclusion_zone_added(nx, ny, nw, nh)

            self._draw_start = None
            self._draw_current = None
            self.image_panel.Refresh()

    def _on_image_motion(self, event: wx.MouseEvent) -> None:
        """Handle mouse motion on image panel."""
        if self._drawing_mode and event.Dragging() and event.LeftIsDown():
            self._draw_current = event.GetPosition()
            self.image_panel.Refresh()

    def _notify_exclusion_zone_added(self, x: float, y: float, w: float, h: float) -> None:
        """Notify parent that an exclusion zone was drawn."""
        # Find MainWindow ancestor
        parent = self.GetParent()
        while parent:
            if hasattr(parent, 'on_exclusion_zone_drawn'):
                parent.on_exclusion_zone_drawn(self.side, self.page_index, x, y, w, h)
                return
            parent = parent.GetParent()

    def _update_scaled_bitmap(self) -> None:
        """Update the displayed bitmap scaled to current size."""
        # Choose which bitmap to display: highlight or original
        bitmap_to_use = self._original_bitmap
        if self._show_diff and self._highlight_bitmap is not None:
            bitmap_to_use = self._highlight_bitmap

        if bitmap_to_use is None:
            return

        # Get available width from parent panel
        parent = self.GetParent()
        if parent is None:
            return

        parent_width = parent.GetClientSize().width
        # Leave margin for scrollbar and padding
        target_w = max(100, parent_width - 50)

        img = bitmap_to_use.ConvertToImage()
        orig_w, orig_h = img.GetWidth(), img.GetHeight()

        # Calculate scale to fit width, maintaining aspect ratio
        scale = target_w / orig_w
        new_w = int(orig_w * scale)
        new_h = int(orig_h * scale)

        img = img.Scale(new_w, new_h, wx.IMAGE_QUALITY_HIGH)
        self._scaled_bitmap = wx.Bitmap(img)

        # Update image panel size and trigger repaint
        self.image_panel.SetMinSize(wx.Size(new_w, new_h))
        self.image_panel.Refresh()

        # Update panel size
        self.SetMinSize(wx.Size(new_w + 20, new_h + 40))
        self.Layout()
        self.Refresh()

    def _on_resize(self, event: wx.SizeEvent) -> None:
        """Handle resize event."""
        event.Skip()
        # Delay update to avoid excessive redraws
        if self._original_bitmap is not None or self._highlight_bitmap is not None:
            wx.CallAfter(self._update_scaled_bitmap)

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
            self.SetBackgroundColour(wx.Colour(204, 229, 255))  # Light blue
        elif self._match_status == MatchStatus.UNMATCHED_LEFT or self._match_status == MatchStatus.UNMATCHED_RIGHT:
            # Orange background for unmatched slides (only on one side)
            self.SetBackgroundColour(wx.Colour(255, 224, 178))  # Light orange
        else:
            self.SetBackgroundColour(wx.WHITE)
        self.Refresh()

    def _on_click(self, event: wx.MouseEvent) -> None:
        """Handle click event."""
        evt = wx.CommandEvent(wx.EVT_BUTTON.typeId)
        evt.SetInt(self.page_index)
        evt.SetString(self.side)
        wx.PostEvent(self.GetParent().GetParent(), evt)

    def _on_double_click(self, event: wx.MouseEvent) -> None:
        """Handle double click event."""
        # Custom event for double click
        pass

    def get_center(self) -> wx.Point:
        """Get center position relative to parent."""
        rect = self.GetRect()
        return wx.Point(rect.x + rect.width // 2, rect.y + rect.height // 2)


class DocumentPanel(scrolled.ScrolledPanel):
    """Scrollable panel containing document page thumbnails."""

    def __init__(self, parent: wx.Window, side: str):
        super().__init__(parent, style=wx.BORDER_SUNKEN)
        self.side = side
        self.document: Optional[Document] = None
        self.thumbnails: List[PageThumbnailPanel] = []
        self._selected_index: Optional[int] = None

        self.SetupScrolling(scroll_x=False, scroll_y=True)
        self.SetBackgroundColour(wx.Colour(248, 249, 250))

        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Header for file info
        self.header_panel = wx.Panel(self)
        self.header_panel.SetBackgroundColour(wx.Colour(52, 58, 64))
        header_sizer = wx.BoxSizer(wx.VERTICAL)

        self.file_name_label = wx.StaticText(self.header_panel, label=f"{'Left' if side == 'left' else 'Right'} Document")
        self.file_name_label.SetForegroundColour(wx.WHITE)
        font = self.file_name_label.GetFont()
        font.SetPointSize(11)
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        self.file_name_label.SetFont(font)

        self.file_path_label = wx.StaticText(self.header_panel, label="No file loaded")
        self.file_path_label.SetForegroundColour(wx.Colour(173, 181, 189))
        font2 = self.file_path_label.GetFont()
        font2.SetPointSize(8)
        self.file_path_label.SetFont(font2)

        header_sizer.Add(self.file_name_label, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        header_sizer.Add(self.file_path_label, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)
        self.header_panel.SetSizer(header_sizer)

        self.sizer.Add(self.header_panel, 0, wx.EXPAND)

        # Placeholder
        self.placeholder = wx.StaticText(
            self,
            label=f"Drop {'Left' if side == 'left' else 'Right'} document here\n(PDF or PowerPoint)",
            style=wx.ALIGN_CENTER
        )
        self.placeholder.SetForegroundColour(wx.Colour(102, 102, 102))
        font = self.placeholder.GetFont()
        font.SetPointSize(12)
        self.placeholder.SetFont(font)
        self.sizer.Add(self.placeholder, 1, wx.ALL | wx.EXPAND, 40)

        self.SetSizer(self.sizer)

        # Setup drop target
        self.SetDropTarget(FileDropTarget(self))

        # Bind resize event
        self.Bind(wx.EVT_SIZE, self._on_panel_resize)

    def _on_panel_resize(self, event: wx.SizeEvent) -> None:
        """Handle panel resize - update all thumbnails."""
        event.Skip()
        wx.CallAfter(self._update_all_thumbnails)

    def _update_all_thumbnails(self) -> None:
        """Update all thumbnail sizes."""
        for thumb in self.thumbnails:
            if thumb._original_bitmap is not None or thumb._highlight_bitmap is not None:
                thumb._update_scaled_bitmap()
        self.Layout()
        self.SetupScrolling(scroll_x=False, scroll_y=True)

    def set_document(self, document: Document) -> None:
        """Set the document to display."""
        self.document = document
        self._update_header()
        self._rebuild_thumbnails()

    def _update_header(self) -> None:
        """Update header with file information."""
        if self.document is None:
            self.file_name_label.SetLabel(f"{'Left' if self.side == 'left' else 'Right'} Document")
            self.file_path_label.SetLabel("No file loaded")
        else:
            self.file_name_label.SetLabel(self.document.name)
            # Truncate path if too long
            path_str = str(self.document.path.parent)
            if len(path_str) > 50:
                path_str = "..." + path_str[-47:]
            self.file_path_label.SetLabel(f"{path_str}  ({self.document.page_count} pages)")
        self.header_panel.Layout()

    def _rebuild_thumbnails(self) -> None:
        """Rebuild thumbnail widgets from document."""

        # Clear existing
        for thumb in self.thumbnails:
            thumb.Destroy()
        self.thumbnails.clear()

        if self.document is None:
            self.placeholder.Show()
            return

        self.placeholder.Hide()

        # Create thumbnails
        for page in self.document.pages:
            thumb = PageThumbnailPanel(self, page.index, self.side)

            if page.thumbnail:
                try:
                    bitmap = pil_to_wxbitmap(page.thumbnail)
                    thumb.set_bitmap(bitmap)
                except Exception:
                    pass  # Silently ignore bitmap conversion errors

            self.thumbnails.append(thumb)
            self.sizer.Add(thumb, 0, wx.ALL | wx.EXPAND, 5)

        self.SetupScrolling(scroll_x=False, scroll_y=True)
        self.Layout()
        self.Refresh()

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

    def get_thumbnail_position(self, index: int) -> Optional[wx.Point]:
        """Get position of a thumbnail for drawing links."""
        if index < len(self.thumbnails):
            thumb = self.thumbnails[index]
            pos = thumb.GetPosition()
            size = thumb.GetSize()
            if self.side == "left":
                # Right edge center
                return self.ClientToScreen(wx.Point(pos.x + size.width, pos.y + size.height // 2))
            else:
                # Left edge center
                return self.ClientToScreen(wx.Point(pos.x, pos.y + size.height // 2))
        return None

    def set_diff_result(self, page_index: int, diff_result: DiffResult, highlight_bitmap: Optional[wx.Bitmap] = None) -> None:
        """Set diff result for a specific page thumbnail."""
        if page_index < len(self.thumbnails):
            self.thumbnails[page_index].set_diff_result(diff_result, highlight_bitmap)

    def scroll_to_page(self, page_index: int) -> None:
        """Scroll to make a specific page visible."""
        if page_index < 0 or page_index >= len(self.thumbnails):
            return

        thumb = self.thumbnails[page_index]
        # Get thumbnail position relative to the scrolled panel
        thumb_pos = thumb.GetPosition()
        thumb_size = thumb.GetSize()

        # Get scroll unit and current position
        scroll_unit = self.GetScrollPixelsPerUnit()[1]
        if scroll_unit == 0:
            scroll_unit = 10

        # Calculate target scroll position (center the thumbnail if possible)
        panel_height = self.GetClientSize().height
        target_y = thumb_pos.y - (panel_height // 2) + (thumb_size.height // 2)
        target_y = max(0, target_y)

        # Scroll to position
        self.Scroll(-1, target_y // scroll_unit)

    def set_drawing_mode(self, enabled: bool) -> None:
        """Enable or disable exclusion zone drawing mode on all thumbnails."""
        for thumb in self.thumbnails:
            thumb.set_drawing_mode(enabled)

    def update_exclusion_zones(self, zones: List[ExclusionZone]) -> None:
        """Update exclusion zone overlays on all thumbnails."""
        # Convert zones to normalized coordinates for display
        zone_tuples = [(z.x, z.y, z.width, z.height) for z in zones if z.enabled]
        for thumb in self.thumbnails:
            # Filter zones by side
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


class FileDropTarget(wx.FileDropTarget):
    """Drop target for document files."""

    def __init__(self, panel: DocumentPanel):
        super().__init__()
        self.panel = panel

    def OnDropFiles(self, x: int, y: int, filenames: List[str]) -> bool:
        """Handle dropped files."""
        valid_files = [
            f for f in filenames
            if f.lower().endswith(('.pdf', '.pptx', '.ppt'))
        ]

        if not valid_files:
            return False

        frame = wx.GetTopLevelParent(self.panel)
        if not hasattr(frame, 'on_file_dropped'):
            return False

        if len(valid_files) >= 2:
            # Two or more files: load first to left, second to right with deferred load
            frame.on_file_dropped(valid_files[0], "left")
            # Defer second document load to allow UI to update
            wx.CallLater(100, frame.on_file_dropped, valid_files[1], "right")
        else:
            # Single file: load to this panel's side
            frame.on_file_dropped(valid_files[0], self.panel.side)

        return True


class LinkOverlayPanel(wx.Panel):
    """Panel that draws lines between matched pages."""

    def __init__(self, parent: wx.Window):
        super().__init__(parent)
        self.links: List[tuple] = []
        self.SetBackgroundColour(wx.Colour(248, 249, 250))

        self.Bind(wx.EVT_PAINT, self._on_paint)

    def set_links(self, links: List[tuple]) -> None:
        """Set links to draw."""
        self.links = links
        self.Refresh()

    def _on_paint(self, event: wx.PaintEvent) -> None:
        """Paint the link lines."""
        dc = wx.PaintDC(self)
        dc.SetBackground(wx.Brush(wx.Colour(248, 249, 250)))
        dc.Clear()

        gc = wx.GraphicsContext.Create(dc)
        if gc is None:
            return

        panel_width = self.GetSize().width
        panel_height = self.GetSize().height

        for left_pos, right_pos, status, has_diff in self.links:
            if left_pos is None or right_pos is None:
                continue

            # Convert screen coordinates to local panel coordinates
            left_local = self.ScreenToClient(left_pos)
            right_local = self.ScreenToClient(right_pos)

            # Only draw if at least one end is visible
            if (left_local.y < 0 and right_local.y < 0) or \
               (left_local.y > panel_height and right_local.y > panel_height):
                continue

            # Clamp X coordinates to panel edges
            left_x = 0  # Left edge
            right_x = panel_width  # Right edge

            # Choose color based on has_diff (True = has differences, False = identical)
            if status == MatchStatus.MATCHED:
                if has_diff:
                    color = wx.Colour(220, 53, 69)  # Red - has differences
                else:
                    color = wx.Colour(40, 167, 69)  # Green - no difference
            else:
                color = wx.Colour(108, 117, 125)  # Gray

            pen = gc.CreatePen(wx.GraphicsPenInfo(color).Width(8))
            gc.SetPen(pen)
            gc.StrokeLine(left_x, left_local.y, right_x, right_local.y)


class ExclusionZoneDialog(wx.Dialog):
    """Dialog for managing exclusion zones."""

    def __init__(self, parent: wx.Window, zone_set: ExclusionZoneSet):
        super().__init__(parent, title="Exclusion Zones", size=(500, 400))
        self.zone_set = zone_set

        self._setup_ui()
        self._populate_list()
        self.Centre()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Zone list
        list_label = wx.StaticText(self, label="Exclusion Zones:")
        main_sizer.Add(list_label, 0, wx.ALL, 5)

        self.zone_list = wx.CheckListBox(self, style=wx.LB_SINGLE)
        main_sizer.Add(self.zone_list, 1, wx.EXPAND | wx.ALL, 5)

        # Preset buttons
        preset_sizer = wx.BoxSizer(wx.HORIZONTAL)
        preset_label = wx.StaticText(self, label="Add Preset:")
        preset_sizer.Add(preset_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        presets = [
            ("Page Number (Bottom)", ExclusionZoneSet.preset_page_number_bottom),
            ("Page Number (Bottom Right)", ExclusionZoneSet.preset_page_number_bottom_right),
            ("Header", ExclusionZoneSet.preset_header),
            ("Footer", ExclusionZoneSet.preset_footer),
            ("Slide Number", ExclusionZoneSet.preset_slide_number_ppt),
        ]

        for name, factory in presets:
            btn = wx.Button(self, label=name, size=(110, -1))
            btn.Bind(wx.EVT_BUTTON, lambda e, f=factory: self._add_preset(f))
            preset_sizer.Add(btn, 0, wx.RIGHT, 3)

        main_sizer.Add(preset_sizer, 0, wx.ALL, 5)

        # Manual coordinate input
        manual_sizer = wx.StaticBoxSizer(wx.StaticBox(self, label="Add Custom Zone"), wx.VERTICAL)

        coord_sizer = wx.FlexGridSizer(rows=2, cols=4, hgap=10, vgap=5)

        coord_sizer.Add(wx.StaticText(self, label="X (0-100%):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.x_input = wx.SpinCtrl(self, min=0, max=100, initial=0, size=(70, -1))
        coord_sizer.Add(self.x_input, 0)

        coord_sizer.Add(wx.StaticText(self, label="Y (0-100%):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.y_input = wx.SpinCtrl(self, min=0, max=100, initial=0, size=(70, -1))
        coord_sizer.Add(self.y_input, 0)

        coord_sizer.Add(wx.StaticText(self, label="Width:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.w_input = wx.SpinCtrl(self, min=1, max=100, initial=20, size=(70, -1))
        coord_sizer.Add(self.w_input, 0)

        coord_sizer.Add(wx.StaticText(self, label="Height:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.h_input = wx.SpinCtrl(self, min=1, max=100, initial=10, size=(70, -1))
        coord_sizer.Add(self.h_input, 0)

        manual_sizer.Add(coord_sizer, 0, wx.ALL, 5)

        # Name and applies_to
        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(self, label="Name:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.name_input = wx.TextCtrl(self, size=(150, -1))
        name_sizer.Add(self.name_input, 0, wx.RIGHT, 10)

        name_sizer.Add(wx.StaticText(self, label="Applies to:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.applies_choice = wx.Choice(self, choices=["Both", "Left Only", "Right Only"])
        self.applies_choice.SetSelection(0)
        name_sizer.Add(self.applies_choice, 0)

        manual_sizer.Add(name_sizer, 0, wx.ALL, 5)

        add_btn = wx.Button(self, label="Add Zone")
        add_btn.Bind(wx.EVT_BUTTON, self._add_custom_zone)
        manual_sizer.Add(add_btn, 0, wx.ALL, 5)

        main_sizer.Add(manual_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Remove button
        remove_btn = wx.Button(self, label="Remove Selected Zone")
        remove_btn.Bind(wx.EVT_BUTTON, self._remove_zone)
        main_sizer.Add(remove_btn, 0, wx.ALL, 5)

        # OK/Cancel buttons
        btn_sizer = self.CreateStdDialogButtonSizer(wx.OK | wx.CANCEL)
        main_sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(main_sizer)

        # Bind checkbox events
        self.zone_list.Bind(wx.EVT_CHECKLISTBOX, self._on_check)

    def _populate_list(self) -> None:
        """Populate the list with current zones."""
        self.zone_list.Clear()
        for zone in self.zone_set.zones:
            label = f"{zone.name or 'Unnamed'} ({zone.x:.0%}, {zone.y:.0%}, {zone.width:.0%}x{zone.height:.0%})"
            idx = self.zone_list.Append(label)
            self.zone_list.Check(idx, zone.enabled)

    def _add_preset(self, factory) -> None:
        """Add a preset zone."""
        zone = factory()
        self.zone_set.add(zone)
        self._populate_list()

    def _add_custom_zone(self, event: wx.Event) -> None:
        """Add a custom zone from inputs."""
        x = self.x_input.GetValue() / 100.0
        y = self.y_input.GetValue() / 100.0
        w = self.w_input.GetValue() / 100.0
        h = self.h_input.GetValue() / 100.0
        name = self.name_input.GetValue() or "Custom"

        applies_idx = self.applies_choice.GetSelection()
        applies_map = {0: AppliesTo.BOTH, 1: AppliesTo.LEFT, 2: AppliesTo.RIGHT}
        applies_to = applies_map[applies_idx]

        zone = ExclusionZone(x=x, y=y, width=w, height=h, name=name, applies_to=applies_to)
        self.zone_set.add(zone)
        self._populate_list()

    def _remove_zone(self, event: wx.Event) -> None:
        """Remove the selected zone."""
        sel = self.zone_list.GetSelection()
        if sel != wx.NOT_FOUND and sel < len(self.zone_set.zones):
            zone = self.zone_set.zones[sel]
            self.zone_set.remove(zone)
            self._populate_list()

    def _on_check(self, event: wx.CommandEvent) -> None:
        """Handle checkbox toggle."""
        idx = event.GetInt()
        if idx < len(self.zone_set.zones):
            self.zone_set.zones[idx].enabled = self.zone_list.IsChecked(idx)


class DiffSummaryDialog(wx.Frame):
    """Modeless dialog showing pages with differences."""

    def __init__(self, parent: wx.Window):
        super().__init__(
            parent,
            title="差分一覧",
            size=(400, 500),
            style=wx.DEFAULT_FRAME_STYLE | wx.FRAME_FLOAT_ON_PARENT
        )
        self.parent = parent
        self.diff_pages: List[tuple] = []  # [(left_idx, right_idx, has_diff), ...]

        self._setup_ui()
        self.Centre()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Header
        header = wx.StaticText(panel, label="差分があるページ一覧")
        header.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        main_sizer.Add(header, 0, wx.ALL, 10)

        # List control
        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "Left Page", width=80)
        self.list_ctrl.InsertColumn(1, "Right Page", width=80)
        self.list_ctrl.InsertColumn(2, "Status", width=100)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self._on_item_activated)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, self._on_item_selected)
        main_sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.ALL, 10)

        # Summary
        self.summary_label = wx.StaticText(panel, label="")
        main_sizer.Add(self.summary_label, 0, wx.ALL, 10)

        # Buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        jump_btn = wx.Button(panel, label="ジャンプ")
        jump_btn.Bind(wx.EVT_BUTTON, self._on_jump)
        btn_sizer.Add(jump_btn, 0, wx.RIGHT, 5)

        export_btn = wx.Button(panel, label="CSVエクスポート")
        export_btn.Bind(wx.EVT_BUTTON, self._on_export)
        btn_sizer.Add(export_btn, 0)

        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 10)

        panel.SetSizer(main_sizer)

    def update_diff_list(self, matching_result: Optional[MatchingResult], diff_scores: dict) -> None:
        """Update the list with current diff information."""
        self.list_ctrl.DeleteAllItems()
        self.diff_pages = []

        if not matching_result:
            self.summary_label.SetLabel("比較結果がありません")
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
                else:
                    identical_count += 1
                    status = "同一"

                self.diff_pages.append((match.left_index, match.right_index, has_diff))
                idx = self.list_ctrl.InsertItem(self.list_ctrl.GetItemCount(), str(match.left_index + 1))
                self.list_ctrl.SetItem(idx, 1, str(match.right_index + 1))
                self.list_ctrl.SetItem(idx, 2, status)

                # Color code
                if has_diff:
                    self.list_ctrl.SetItemBackgroundColour(idx, wx.Colour(255, 200, 200))
                else:
                    self.list_ctrl.SetItemBackgroundColour(idx, wx.Colour(200, 255, 200))
            else:
                unmatched_count += 1

        self.summary_label.SetLabel(
            f"差分あり: {diff_count}  同一: {identical_count}  未マッチ: {unmatched_count}"
        )

    def _on_item_activated(self, event: wx.ListEvent) -> None:
        """Handle double click on item."""
        self._jump_to_selected()

    def _on_item_selected(self, event: wx.ListEvent) -> None:
        """Handle item selection."""
        pass

    def _on_jump(self, event: wx.Event) -> None:
        """Jump to selected page."""
        self._jump_to_selected()

    def _jump_to_selected(self) -> None:
        """Jump to the selected page in main window."""
        sel = self.list_ctrl.GetFirstSelected()
        if sel != wx.NOT_FOUND and sel < len(self.diff_pages):
            left_idx, right_idx, _ = self.diff_pages[sel]
            if hasattr(self.parent, 'jump_to_page_pair'):
                self.parent.jump_to_page_pair(left_idx, right_idx)

    def _on_export(self, event: wx.Event) -> None:
        """Export diff list to CSV."""
        with wx.FileDialog(
            self,
            "Export CSV",
            wildcard="CSV Files (*.csv)|*.csv",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        ) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return

            path = dialog.GetPath()
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write("Left Page,Right Page,Status\n")
                    for left_idx, right_idx, has_diff in self.diff_pages:
                        status = "差分あり" if has_diff else "同一"
                        f.write(f"{left_idx + 1},{right_idx + 1},{status}\n")
                wx.MessageBox(f"エクスポート完了: {path}", "完了", wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"エクスポート失敗: {e}", "エラー", wx.OK | wx.ICON_ERROR)

    def get_diff_page_indices(self) -> List[tuple]:
        """Get list of (left_idx, right_idx) for pages with differences."""
        return [(left, right) for left, right, has_diff in self.diff_pages if has_diff]


class MainWindow(wx.Frame):
    """Main application window."""

    def __init__(self):
        super().__init__(
            None,
            title="PPT/PDF Comparator (wxPython)",
            size=(1600, 1000)
        )
        self.SetMinSize((800, 600))
        self.Maximize(True)  # 起動時に最大化

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

        # Performance: batch updates
        self._update_pending = False

        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()

        # Accept drops on frame
        self.SetDropTarget(MainDropTarget(self))

        self.Centre()

        # Bind close event
        self.Bind(wx.EVT_CLOSE, self._on_close)

    def _setup_ui(self) -> None:
        """Set up the main UI layout."""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Main content area (horizontal)
        content_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left document panel
        self.left_panel = DocumentPanel(panel, "left")
        self.left_panel.SetMinSize((200, -1))

        # Center link overlay panel
        self.link_overlay = LinkOverlayPanel(panel)
        self.link_overlay.SetMinSize((80, -1))

        # Right document panel
        self.right_panel = DocumentPanel(panel, "right")
        self.right_panel.SetMinSize((200, -1))

        content_sizer.Add(self.left_panel, 1, wx.EXPAND | wx.ALL, 5)
        content_sizer.Add(self.link_overlay, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 5)
        content_sizer.Add(self.right_panel, 1, wx.EXPAND | wx.ALL, 5)

        main_sizer.Add(content_sizer, 1, wx.EXPAND)

        # Bottom sync scroll bar area
        sync_panel = wx.Panel(panel)
        sync_panel.SetBackgroundColour(wx.Colour(220, 220, 220))
        sync_panel.SetMinSize((-1, 40))
        sync_sizer = wx.BoxSizer(wx.HORIZONTAL)

        sync_label = wx.StaticText(sync_panel, label="連動スクロール:")
        sync_label.SetForegroundColour(wx.Colour(80, 80, 80))
        sync_sizer.Add(sync_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)

        self.sync_scrollbar = wx.Slider(
            sync_panel,
            value=0,
            minValue=0,
            maxValue=100,
            style=wx.SL_HORIZONTAL
        )
        sync_sizer.Add(self.sync_scrollbar, 1, wx.EXPAND | wx.ALL, 5)

        # Sync toggle checkbox
        self.sync_checkbox = wx.CheckBox(sync_panel, label="連動")
        self.sync_checkbox.SetValue(False)
        sync_sizer.Add(self.sync_checkbox, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        sync_panel.SetSizer(sync_sizer)
        main_sizer.Add(sync_panel, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        panel.SetSizer(main_sizer)

        # Bind custom event for page clicks
        self.Bind(wx.EVT_BUTTON, self._on_page_clicked)

        # Bind scroll events to update links
        self.left_panel.Bind(wx.EVT_SCROLLWIN, self._on_panel_scroll)
        self.right_panel.Bind(wx.EVT_SCROLLWIN, self._on_panel_scroll)
        self.Bind(wx.EVT_SIZE, self._on_window_resize)

        # Bind sync scrollbar
        self.sync_scrollbar.Bind(wx.EVT_SLIDER, self._on_sync_scroll)
        self.sync_checkbox.Bind(wx.EVT_CHECKBOX, self._on_sync_toggle)

    def _setup_menu(self) -> None:
        """Set up the menu bar."""
        menubar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()
        open_left = file_menu.Append(wx.ID_ANY, "Open Left...\tCtrl+O")
        open_right = file_menu.Append(wx.ID_ANY, "Open Right...\tCtrl+Shift+O")
        file_menu.AppendSeparator()
        save_session = file_menu.Append(wx.ID_ANY, "Save Session...\tCtrl+S")
        load_session = file_menu.Append(wx.ID_ANY, "Load Session...\tCtrl+L")
        file_menu.AppendSeparator()
        exit_item = file_menu.Append(wx.ID_EXIT, "Exit\tCtrl+Q")

        self.Bind(wx.EVT_MENU, self._open_left_file, open_left)
        self.Bind(wx.EVT_MENU, self._open_right_file, open_right)
        self.Bind(wx.EVT_MENU, self._save_session, save_session)
        self.Bind(wx.EVT_MENU, self._load_session, load_session)
        self.Bind(wx.EVT_MENU, lambda e: self.Close(), exit_item)

        menubar.Append(file_menu, "File")

        # Compare menu
        compare_menu = wx.Menu()
        run_compare = compare_menu.Append(wx.ID_ANY, "Run Comparison\tF5")
        compare_menu.AppendSeparator()
        exclusion_zones = compare_menu.Append(wx.ID_ANY, "Exclusion Zones...")
        compare_menu.AppendSeparator()
        clear_links = compare_menu.Append(wx.ID_ANY, "Clear Manual Links")

        self.Bind(wx.EVT_MENU, self._run_comparison, run_compare)
        self.Bind(wx.EVT_MENU, self._show_exclusion_zones_dialog, exclusion_zones)
        self.Bind(wx.EVT_MENU, self._clear_manual_links, clear_links)

        menubar.Append(compare_menu, "Compare")

        # Export menu
        export_menu = wx.Menu()
        export_pdf = export_menu.Append(wx.ID_ANY, "Export to PDF...")
        export_html = export_menu.Append(wx.ID_ANY, "Export to HTML...")

        self.Bind(wx.EVT_MENU, self._export_pdf, export_pdf)
        self.Bind(wx.EVT_MENU, self._export_html, export_html)

        menubar.Append(export_menu, "Export")

        self.SetMenuBar(menubar)

    def _setup_toolbar(self) -> None:
        """Set up the toolbar."""
        toolbar = self.CreateToolBar()

        open_left_btn = toolbar.AddTool(wx.ID_ANY, "Open Left", wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN))
        open_right_btn = toolbar.AddTool(wx.ID_ANY, "Open Right", wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN))
        toolbar.AddSeparator()
        compare_btn = toolbar.AddTool(wx.ID_ANY, "Compare", wx.ArtProvider.GetBitmap(wx.ART_FIND))
        toolbar.AddSeparator()

        # Diff navigation buttons
        prev_diff_btn = toolbar.AddTool(wx.ID_ANY, "Prev Diff", wx.ArtProvider.GetBitmap(wx.ART_GO_UP), shortHelp="前の差分 (Ctrl+Up)")
        next_diff_btn = toolbar.AddTool(wx.ID_ANY, "Next Diff", wx.ArtProvider.GetBitmap(wx.ART_GO_DOWN), shortHelp="次の差分 (Ctrl+Down)")
        toolbar.AddSeparator()

        # Diff summary button
        diff_summary_btn = toolbar.AddTool(wx.ID_ANY, "Diff List", wx.ArtProvider.GetBitmap(wx.ART_LIST_VIEW), shortHelp="差分一覧")
        toolbar.AddSeparator()

        # Color picker button
        color_btn = toolbar.AddTool(wx.ID_ANY, "Color", wx.ArtProvider.GetBitmap(wx.ART_HELP_SETTINGS), shortHelp="ハイライト色")
        toolbar.AddSeparator()

        # Exclusion zone drawing mode toggle
        self._draw_mode_btn = toolbar.AddCheckTool(
            wx.ID_ANY, "Draw Zone",
            wx.ArtProvider.GetBitmap(wx.ART_CUT),
            shortHelp="除外領域描画モード"
        )

        # Remove zone buttons
        remove_zone_btn = toolbar.AddTool(wx.ID_ANY, "Remove Zone", wx.ArtProvider.GetBitmap(wx.ART_DELETE), shortHelp="最後の除外領域を削除")
        clear_zones_btn = toolbar.AddTool(wx.ID_ANY, "Clear Zones", wx.ArtProvider.GetBitmap(wx.ART_CROSS_MARK), shortHelp="全ての除外領域を削除")
        toolbar.AddSeparator()

        export_pdf_btn = toolbar.AddTool(wx.ID_ANY, "Export PDF", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))
        export_html_btn = toolbar.AddTool(wx.ID_ANY, "Export HTML", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE_AS))

        self.Bind(wx.EVT_TOOL, self._open_left_file, open_left_btn)
        self.Bind(wx.EVT_TOOL, self._open_right_file, open_right_btn)
        self.Bind(wx.EVT_TOOL, self._run_comparison, compare_btn)
        self.Bind(wx.EVT_TOOL, self._go_prev_diff, prev_diff_btn)
        self.Bind(wx.EVT_TOOL, self._go_next_diff, next_diff_btn)
        self.Bind(wx.EVT_TOOL, self._show_diff_summary, diff_summary_btn)
        self.Bind(wx.EVT_TOOL, self._pick_highlight_color, color_btn)
        self.Bind(wx.EVT_TOOL, self._toggle_drawing_mode, self._draw_mode_btn)
        self.Bind(wx.EVT_TOOL, self._remove_last_exclusion_zone, remove_zone_btn)
        self.Bind(wx.EVT_TOOL, self._clear_all_exclusion_zones, clear_zones_btn)
        self.Bind(wx.EVT_TOOL, self._export_pdf, export_pdf_btn)
        self.Bind(wx.EVT_TOOL, self._export_html, export_html_btn)

        toolbar.Realize()

        # Keyboard shortcuts for navigation
        accel_tbl = wx.AcceleratorTable([
            (wx.ACCEL_CTRL, wx.WXK_UP, prev_diff_btn.GetId()),
            (wx.ACCEL_CTRL, wx.WXK_DOWN, next_diff_btn.GetId()),
        ])
        self.SetAcceleratorTable(accel_tbl)

    def _setup_statusbar(self) -> None:
        """Set up the status bar."""
        self.statusbar = self.CreateStatusBar()
        self.statusbar.SetStatusText("Ready - Drop files to begin")

    def _open_file_dialog(self) -> Optional[str]:
        """Open file dialog and return selected path."""
        with wx.FileDialog(
            self,
            "Open Document",
            wildcard="Documents (*.pdf;*.pptx;*.ppt)|*.pdf;*.pptx;*.ppt|PDF Files (*.pdf)|*.pdf|PowerPoint Files (*.pptx;*.ppt)|*.pptx;*.ppt",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return None
            return dialog.GetPath()

    def _open_left_file(self, event: wx.Event) -> None:
        """Open file dialog for left document."""
        path = self._open_file_dialog()
        if path:
            self._load_document(path, "left")

    def _open_right_file(self, event: wx.Event) -> None:
        """Open file dialog for right document."""
        path = self._open_file_dialog()
        if path:
            self._load_document(path, "right")

    def on_file_dropped(self, path: str, side: str) -> None:
        """Handle file dropped on a panel."""
        self._load_document(path, side)

    def _load_document(self, path: str, side: str) -> None:
        """Load a document file."""
        try:
            doc = Document.from_file(path)

            # Show progress dialog
            progress = wx.ProgressDialog(
                "Loading",
                f"Loading {Path(path).name}...",
                maximum=100,
                parent=self,
                style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE
            )

            def progress_callback(current, total):
                progress.Update(int(current / total * 100))
                wx.GetApp().Yield()

            doc.load(progress_callback=progress_callback)
            progress.Destroy()

            if side == "left":
                self.left_doc = doc
                self.left_panel.set_document(doc)
                self.session.left_document_path = path
            else:
                self.right_doc = doc
                self.right_panel.set_document(doc)
                self.session.right_document_path = path

            self._update_status()

            # 両方のドキュメントが読み込まれたら自動的に比較を実行
            if self.left_doc and self.right_doc:
                wx.CallAfter(self._run_comparison, None)

        except Exception as e:
            wx.MessageBox(
                f"Failed to load document:\n{str(e)}",
                "Error",
                wx.OK | wx.ICON_ERROR
            )

    def _run_comparison(self, event: wx.Event) -> None:
        """Run page matching comparison."""
        if not self.left_doc or not self.right_doc:
            wx.MessageBox(
                "Please load both documents before comparing.",
                "Warning",
                wx.OK | wx.ICON_WARNING
            )
            return

        progress = wx.ProgressDialog(
            "Comparing",
            "Matching pages...",
            maximum=100,
            parent=self,
            style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE
        )

        def progress_callback(current, total, msg=""):
            progress.Update(int(current / total * 100) if total > 0 else 0, msg or "Matching pages...")
            wx.GetApp().Yield()

        # Phase 1: Match pages
        matcher = PageMatcher()
        self.matching_result = matcher.match(
            self.left_doc, self.right_doc,
            progress_callback=progress_callback
        )

        # Phase 2: Compute differences for matched pairs
        progress.Update(0, "Computing differences...")
        comparator = ImageComparator(highlight_color=self._highlight_color)
        matched_pairs = self.matching_result.get_matched_pairs()
        total_pairs = len(matched_pairs)
        self.diff_scores = {}  # Clear previous scores

        for i, (left_idx, right_idx, score) in enumerate(matched_pairs):
            progress.Update(int((i + 1) / total_pairs * 100) if total_pairs > 0 else 100,
                          f"Computing diff {i + 1}/{total_pairs}...")
            wx.GetApp().Yield()

            # Get images for comparison
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
                    # Convert PIL Image to wx.Bitmap
                    left_highlight_bitmap = pil_to_wxbitmap(diff_result.highlight_image)
                    self.left_panel.set_diff_result(left_idx, diff_result, left_highlight_bitmap)

                    # Also create highlight for right side (compare right to left)
                    diff_result_right = comparator.compare(
                        right_page.thumbnail, left_page.thumbnail,
                        exclusion_zones=all_zones
                    )
                    if diff_result_right.highlight_image:
                        right_highlight_bitmap = pil_to_wxbitmap(diff_result_right.highlight_image)
                        self.right_panel.set_diff_result(right_idx, diff_result_right, right_highlight_bitmap)

        progress.Destroy()

        self.session.matching_result = self.matching_result
        self._current_diff_index = -1  # Reset diff navigation

        # Update UI
        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_status()
        self._update_links()

        # Update diff summary dialog if open
        if self._diff_summary_dialog and self._diff_summary_dialog.IsShown():
            self._diff_summary_dialog.update_diff_list(self.matching_result, self.diff_scores)

    def _on_page_clicked(self, event: wx.CommandEvent) -> None:
        """Handle page click for manual linking and scroll to paired slide."""
        index = event.GetInt()
        side = event.GetString()

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

    def _create_manual_link(self, left_index: int, right_index: int) -> None:
        """Create a manual link between pages."""
        if self.matching_result is None:
            self.matching_result = MatchingResult()

        self.matching_result.set_manual_match(left_index, right_index)
        self.session.matching_result = self.matching_result

        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_links()

        self.statusbar.SetStatusText(
            f"Linked: Left page {left_index + 1} <-> Right page {right_index + 1}"
        )

    def _clear_manual_links(self, event: wx.Event) -> None:
        """Clear all manual links."""
        if self.matching_result:
            self._run_comparison(event)

    def _show_exclusion_zones_dialog(self, event: wx.Event) -> None:
        """Show the exclusion zones dialog."""
        dialog = ExclusionZoneDialog(self, self.exclusion_zones)
        if dialog.ShowModal() == wx.ID_OK:
            # Zones were modified, update session
            self.session.exclusion_zones = self.exclusion_zones
            self.statusbar.SetStatusText(
                f"Exclusion zones updated ({len(self.exclusion_zones.zones)} zones)"
            )
        dialog.Destroy()

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

        self.statusbar.SetStatusText(" | ".join(parts) if parts else "Ready")

    def _save_session(self, event: wx.Event) -> None:
        """Save current session."""
        with wx.FileDialog(
            self,
            "Save Session",
            wildcard="Session Files (*.json)|*.json",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        ) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = dialog.GetPath()
            self.session.save(path)
            self.statusbar.SetStatusText(f"Session saved to {path}")

    def _load_session(self, event: wx.Event) -> None:
        """Load a session file."""
        with wx.FileDialog(
            self,
            "Load Session",
            wildcard="Session Files (*.json)|*.json",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = dialog.GetPath()
            try:
                self.session = Session.load(path)
                if self.session.left_document_path:
                    self._load_document(self.session.left_document_path, "left")
                if self.session.right_document_path:
                    self._load_document(self.session.right_document_path, "right")
                self.matching_result = self.session.matching_result
                self.exclusion_zones = self.session.exclusion_zones
                if self.matching_result:
                    self.left_panel.update_match_status(self.matching_result)
                    self.right_panel.update_match_status(self.matching_result)
                    self._update_links()
                self.statusbar.SetStatusText(f"Session loaded from {path}")
            except Exception as e:
                wx.MessageBox(f"Failed to load session:\n{e}", "Error", wx.OK | wx.ICON_ERROR)

    def _on_panel_scroll(self, event: wx.ScrollWinEvent) -> None:
        """Handle scroll event from document panels."""
        event.Skip()

        # Sync scroll if enabled
        if hasattr(self, 'sync_checkbox') and self.sync_checkbox.GetValue():
            source_panel = event.GetEventObject()
            if source_panel == self.left_panel:
                self._sync_scroll_from(self.left_panel, self.right_panel)
            elif source_panel == self.right_panel:
                self._sync_scroll_from(self.right_panel, self.left_panel)

        wx.CallAfter(self._update_links)
        wx.CallAfter(self._update_sync_scrollbar)

    def _on_window_resize(self, event: wx.SizeEvent) -> None:
        """Handle window resize event."""
        event.Skip()
        wx.CallAfter(self._update_links)

    def _on_sync_scroll(self, event: wx.CommandEvent) -> None:
        """Handle sync scrollbar movement - scroll through matched pairs."""
        if not self.left_doc or not self.right_doc:
            return

        value = self.sync_scrollbar.GetValue()

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

                wx.CallAfter(self._update_links)
                return

        # Fallback: scroll both panels to the same percentage
        percent = value / 100.0
        self._scroll_panel_to_percent(self.left_panel, percent)
        self._scroll_panel_to_percent(self.right_panel, percent)

        wx.CallAfter(self._update_links)

    def _on_sync_toggle(self, event: wx.CommandEvent) -> None:
        """Handle sync checkbox toggle."""
        if self.sync_checkbox.GetValue():
            # Sync both panels when enabled
            wx.CallAfter(self._update_sync_scrollbar)

    def _sync_scroll_from(self, source: DocumentPanel, target: DocumentPanel) -> None:
        """Sync scroll position from source to target panel, trying to align paired slides."""
        if not self.matching_result:
            # Fallback to percentage-based sync if no matching
            source_range = source.GetScrollRange(wx.VERTICAL)
            source_pos = source.GetScrollPos(wx.VERTICAL)
            if source_range > 0:
                percent = source_pos / source_range
                target_range = target.GetScrollRange(wx.VERTICAL)
                target_pos = int(percent * target_range)
                target.Scroll(-1, target_pos)
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

        panel_height = panel.GetClientSize().height
        panel_center_y = panel_height // 2

        # Find which thumbnail is closest to the center
        best_idx = 0
        best_distance = float('inf')

        for i, thumb in enumerate(panel.thumbnails):
            thumb_pos = thumb.GetPosition()
            thumb_size = thumb.GetSize()
            thumb_center_y = thumb_pos.y + thumb_size.height // 2

            distance = abs(thumb_center_y - panel_center_y)
            if distance < best_distance:
                best_distance = distance
                best_idx = i

        return best_idx

    def _scroll_panel_to_percent(self, panel: DocumentPanel, percent: float) -> None:
        """Scroll panel to a percentage position."""
        scroll_range = panel.GetScrollRange(wx.VERTICAL)
        if scroll_range > 0:
            pos = int(percent * scroll_range)
            panel.Scroll(-1, pos)

    def _update_sync_scrollbar(self) -> None:
        """Update sync scrollbar position based on currently visible pair."""
        if not hasattr(self, 'sync_scrollbar'):
            return

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
                            self.sync_scrollbar.SetValue(min(100, max(0, percent)))
                            return

        # Fallback: use scroll percentage
        scroll_range = self.left_panel.GetScrollRange(wx.VERTICAL)
        scroll_pos = self.left_panel.GetScrollPos(wx.VERTICAL)

        if scroll_range > 0:
            percent = int((scroll_pos / scroll_range) * 100)
            self.sync_scrollbar.SetValue(min(100, max(0, percent)))

    def _update_links(self) -> None:
        """Update the link lines between matched pages."""
        if not self.matching_result:
            self.link_overlay.set_links([])
            return

        links = []
        for match in self.matching_result.matches:
            if match.status != MatchStatus.MATCHED:
                continue
            if match.left_index is None or match.right_index is None:
                continue

            # Get positions of thumbnails
            left_pos = self.left_panel.get_thumbnail_position(match.left_index)
            right_pos = self.right_panel.get_thumbnail_position(match.right_index)

            if left_pos and right_pos:
                # Use has_differences for coloring (True = red, False = green)
                has_diff = self.diff_scores.get((match.left_index, match.right_index), False)
                links.append((left_pos, right_pos, match.status, has_diff))

        self.link_overlay.set_links(links)

    def _on_close(self, event: wx.CloseEvent) -> None:
        """Handle window close event."""
        # Close diff summary dialog if open
        if self._diff_summary_dialog:
            self._diff_summary_dialog.Destroy()
            self._diff_summary_dialog = None

        # Close PowerPoint cache if used
        try:
            from ..core.document import close_powerpoint_cache
            close_powerpoint_cache()
        except Exception:
            pass

        event.Skip()

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

    def _go_prev_diff(self, event: wx.Event) -> None:
        """Navigate to previous page with differences."""
        diff_pages = self._get_diff_pages()
        if not diff_pages:
            self.statusbar.SetStatusText("差分がありません")
            return

        if self._current_diff_index <= 0:
            self._current_diff_index = len(diff_pages) - 1
        else:
            self._current_diff_index -= 1

        left_idx, right_idx = diff_pages[self._current_diff_index]
        self.jump_to_page_pair(left_idx, right_idx)
        self.statusbar.SetStatusText(f"差分 {self._current_diff_index + 1}/{len(diff_pages)}")

    def _go_next_diff(self, event: wx.Event) -> None:
        """Navigate to next page with differences."""
        diff_pages = self._get_diff_pages()
        if not diff_pages:
            self.statusbar.SetStatusText("差分がありません")
            return

        if self._current_diff_index >= len(diff_pages) - 1:
            self._current_diff_index = 0
        else:
            self._current_diff_index += 1

        left_idx, right_idx = diff_pages[self._current_diff_index]
        self.jump_to_page_pair(left_idx, right_idx)
        self.statusbar.SetStatusText(f"差分 {self._current_diff_index + 1}/{len(diff_pages)}")

    def jump_to_page_pair(self, left_idx: int, right_idx: int) -> None:
        """Jump to a specific page pair."""
        self.left_panel.scroll_to_page(left_idx)
        self.right_panel.scroll_to_page(right_idx)
        self.left_panel.set_selected(left_idx)
        self.right_panel.set_selected(right_idx)
        wx.CallAfter(self._update_links)

    def _show_diff_summary(self, event: wx.Event) -> None:
        """Show the diff summary dialog."""
        if self._diff_summary_dialog is None:
            self._diff_summary_dialog = DiffSummaryDialog(self)

        self._diff_summary_dialog.update_diff_list(self.matching_result, self.diff_scores)
        self._diff_summary_dialog.Show()
        self._diff_summary_dialog.Raise()

    def _pick_highlight_color(self, event: wx.Event) -> None:
        """Open color picker for highlight color."""
        current = wx.Colour(*self._highlight_color)
        data = wx.ColourData()
        data.SetColour(current)

        dialog = wx.ColourDialog(self, data)
        if dialog.ShowModal() == wx.ID_OK:
            color = dialog.GetColourData().GetColour()
            self._highlight_color = (color.Red(), color.Green(), color.Blue())
            self.statusbar.SetStatusText(f"ハイライト色を変更しました: RGB{self._highlight_color}")

            # Re-run comparison with new color if documents are loaded
            if self.left_doc and self.right_doc and self.matching_result:
                self._run_comparison(None)

        dialog.Destroy()

    def _toggle_drawing_mode(self, event: wx.Event) -> None:
        """Toggle exclusion zone drawing mode."""
        self._drawing_mode = not self._drawing_mode
        self.left_panel.set_drawing_mode(self._drawing_mode)
        self.right_panel.set_drawing_mode(self._drawing_mode)

        if self._drawing_mode:
            self.statusbar.SetStatusText("除外領域描画モード: 画像上をドラッグして除外領域を指定してください")
        else:
            self.statusbar.SetStatusText("除外領域描画モードを終了しました")
            # Update exclusion zone overlays
            self._update_exclusion_zone_overlays()

    def on_exclusion_zone_drawn(self, side: str, page_index: int, x: float, y: float, w: float, h: float) -> None:
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

        # Update overlays
        self._update_exclusion_zone_overlays()

        self.statusbar.SetStatusText(
            f"除外領域を追加しました: {zone.name} ({x*100:.0f}%, {y*100:.0f}%, {w*100:.0f}%x{h*100:.0f}%)"
        )

    def on_exclusion_zone_delete(self, zone_index: int) -> None:
        """Handle exclusion zone delete request."""
        if 0 <= zone_index < len(self.exclusion_zones.zones):
            removed = self.exclusion_zones.zones.pop(zone_index)
            self.session.exclusion_zones = self.exclusion_zones
            self._update_exclusion_zone_overlays()
            self.statusbar.SetStatusText(f"除外領域を削除しました: {removed.name}")

    def _remove_last_exclusion_zone(self, event: wx.Event) -> None:
        """Remove the most recently added exclusion zone."""
        if self.exclusion_zones.zones:
            removed = self.exclusion_zones.zones.pop()
            self.session.exclusion_zones = self.exclusion_zones
            self._update_exclusion_zone_overlays()
            self.statusbar.SetStatusText(f"除外領域を削除しました: {removed.name}")

    def _clear_all_exclusion_zones(self, event: wx.Event) -> None:
        """Clear all exclusion zones."""
        if self.exclusion_zones.zones:
            count = len(self.exclusion_zones.zones)
            self.exclusion_zones.zones.clear()
            self.session.exclusion_zones = self.exclusion_zones
            self._update_exclusion_zone_overlays()
            self.statusbar.SetStatusText(f"全ての除外領域を削除しました ({count} 領域)")

    def _update_exclusion_zone_overlays(self) -> None:
        """Update exclusion zone overlays on both panels."""
        left_zones = self.exclusion_zones.get_zones_for("left")
        right_zones = self.exclusion_zones.get_zones_for("right")
        self.left_panel.update_exclusion_zones(left_zones)
        self.right_panel.update_exclusion_zones(right_zones)

    def _export_pdf(self, event: wx.Event) -> None:
        """Export comparison report to PDF."""
        wx.MessageBox(
            "PDF export will be implemented in the full version.",
            "Export PDF",
            wx.OK | wx.ICON_INFORMATION
        )

    def _export_html(self, event: wx.Event) -> None:
        """Export comparison report to HTML."""
        wx.MessageBox(
            "HTML export will be implemented in the full version.",
            "Export HTML",
            wx.OK | wx.ICON_INFORMATION
        )


class MainDropTarget(wx.FileDropTarget):
    """Drop target for main window."""

    def __init__(self, frame: MainWindow):
        super().__init__()
        self.frame = frame

    def OnDropFiles(self, x: int, y: int, filenames: List[str]) -> bool:
        """Handle dropped files on main window."""
        valid_files = [
            f for f in filenames
            if f.lower().endswith(('.pdf', '.pptx', '.ppt'))
        ]

        if len(valid_files) >= 2:
            self.frame._load_document(valid_files[0], "left")
            # Defer second document load to allow UI to update
            wx.CallLater(100, self.frame._load_document, valid_files[1], "right")
        elif len(valid_files) == 1:
            if self.frame.left_doc is None:
                self.frame._load_document(valid_files[0], "left")
            elif self.frame.right_doc is None:
                self.frame._load_document(valid_files[0], "right")
            else:
                self.frame._load_document(valid_files[0], "left")

        return bool(valid_files)
