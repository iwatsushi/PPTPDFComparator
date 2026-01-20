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

        self.SetBackgroundColour(wx.WHITE)

        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Image display
        self.image_ctrl = wx.StaticBitmap(self)
        self.sizer.Add(self.image_ctrl, 1, wx.ALL | wx.EXPAND, 5)

        # Page label
        self.page_label = wx.StaticText(self, label=f"Page {page_index + 1}")
        self.page_label.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        self.sizer.Add(self.page_label, 0, wx.ALL | wx.CENTER, 2)

        self.SetSizer(self.sizer)

        # Bind events
        self.Bind(wx.EVT_LEFT_DOWN, self._on_click)
        self.Bind(wx.EVT_LEFT_DCLICK, self._on_double_click)
        self.Bind(wx.EVT_SIZE, self._on_resize)
        self.image_ctrl.Bind(wx.EVT_LEFT_DOWN, self._on_click)
        self.image_ctrl.Bind(wx.EVT_LEFT_DCLICK, self._on_double_click)

    def set_bitmap(self, bitmap: wx.Bitmap) -> None:
        """Set the thumbnail image."""
        self._original_bitmap = bitmap
        self._highlight_bitmap = None  # Clear highlight when original changes
        self._update_scaled_bitmap()

    def set_diff_result(self, diff_result: Optional[DiffResult], highlight_image: Optional[wx.Bitmap] = None) -> None:
        """Set the diff result and highlighted image."""
        self._diff_result = diff_result
        self._highlight_bitmap = highlight_image
        print(f"[DEBUG] set_diff_result called for page {self.page_index}: highlight_bitmap={highlight_image is not None}, show_diff={self._show_diff}")

        # 直接image_ctrlに設定
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
                scaled_bitmap = wx.Bitmap(img)
                self.image_ctrl.SetBitmap(scaled_bitmap)
                print(f"[DEBUG] Directly set highlight bitmap for page {self.page_index}")

        self.Layout()
        self.Refresh()
        self.Update()

    def _update_scaled_bitmap(self) -> None:
        """Update the displayed bitmap scaled to current size."""
        # Choose which bitmap to display: highlight or original
        bitmap_to_use = self._original_bitmap
        if self._show_diff and self._highlight_bitmap is not None:
            bitmap_to_use = self._highlight_bitmap
            print(f"[DEBUG] Using highlight bitmap for page {self.page_index}")

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
        scaled_bitmap = wx.Bitmap(img)
        self.image_ctrl.SetBitmap(scaled_bitmap)

        # Update panel size
        self.SetMinSize((new_w + 20, new_h + 40))
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
        elif self._match_status == MatchStatus.UNMATCHED_LEFT:
            self.SetBackgroundColour(wx.Colour(255, 243, 205))  # Light yellow
        elif self._match_status == MatchStatus.UNMATCHED_RIGHT:
            self.SetBackgroundColour(wx.Colour(212, 237, 218))  # Light green
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
        print(f"[DEBUG] _rebuild_thumbnails called, document={self.document}")

        # Clear existing
        for thumb in self.thumbnails:
            thumb.Destroy()
        self.thumbnails.clear()

        if self.document is None:
            self.placeholder.Show()
            return

        self.placeholder.Hide()

        print(f"[DEBUG] Creating {len(self.document.pages)} thumbnails")

        # Create thumbnails
        for page in self.document.pages:
            thumb = PageThumbnailPanel(self, page.index, self.side)

            if page.thumbnail:
                print(f"[DEBUG] Page {page.index}: Converting PIL image {page.thumbnail.size} to bitmap")
                try:
                    bitmap = pil_to_wxbitmap(page.thumbnail)
                    print(f"[DEBUG] Page {page.index}: Bitmap created, size={bitmap.GetWidth()}x{bitmap.GetHeight()}")
                    thumb.set_bitmap(bitmap)
                except Exception as e:
                    print(f"[DEBUG] Page {page.index}: Error converting to bitmap: {e}")
            else:
                print(f"[DEBUG] Page {page.index}: No thumbnail")

            self.thumbnails.append(thumb)
            self.sizer.Add(thumb, 0, wx.ALL | wx.EXPAND, 5)

        self.SetupScrolling(scroll_x=False, scroll_y=True)
        self.Layout()
        self.Refresh()
        print(f"[DEBUG] _rebuild_thumbnails complete, {len(self.thumbnails)} thumbnails created")

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
            # Two or more files: load first to left, second to right
            frame.on_file_dropped(valid_files[0], "left")
            frame.on_file_dropped(valid_files[1], "right")
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

        for left_pos, right_pos, status, similarity in self.links:
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

            # Choose color based on status and similarity
            if status == MatchStatus.MATCHED:
                if similarity >= 0.95:
                    color = wx.Colour(40, 167, 69)  # Green
                elif similarity >= 0.8:
                    color = wx.Colour(255, 193, 7)  # Yellow
                else:
                    color = wx.Colour(220, 53, 69)  # Red
            else:
                color = wx.Colour(108, 117, 125)  # Gray

            pen = gc.CreatePen(wx.GraphicsPenInfo(color).Width(2))
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

        self._selected_left: Optional[int] = None
        self._selected_right: Optional[int] = None

        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()

        # Accept drops on frame
        self.SetDropTarget(MainDropTarget(self))

        self.Centre()

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
        export_pdf_btn = toolbar.AddTool(wx.ID_ANY, "Export PDF", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE))
        export_html_btn = toolbar.AddTool(wx.ID_ANY, "Export HTML", wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE_AS))

        self.Bind(wx.EVT_TOOL, self._open_left_file, open_left_btn)
        self.Bind(wx.EVT_TOOL, self._open_right_file, open_right_btn)
        self.Bind(wx.EVT_TOOL, self._run_comparison, compare_btn)
        self.Bind(wx.EVT_TOOL, self._export_pdf, export_pdf_btn)
        self.Bind(wx.EVT_TOOL, self._export_html, export_html_btn)

        toolbar.Realize()

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
        import traceback
        try:
            doc = Document.from_file(path)
            print(f"[DEBUG] Document created: {doc.path}, type: {doc.doc_type}")

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

            print(f"[DEBUG] Document loaded: {doc.page_count} pages")
            for i, page in enumerate(doc.pages[:3]):  # Show first 3 pages info
                print(f"[DEBUG] Page {i}: thumbnail={page.thumbnail}, size={page.thumbnail.size if page.thumbnail else 'None'}")

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
            traceback.print_exc()
            wx.MessageBox(
                f"Failed to load document:\n{str(e)}\n\n{traceback.format_exc()}",
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
        comparator = ImageComparator()
        matched_pairs = self.matching_result.get_matched_pairs()
        total_pairs = len(matched_pairs)

        for i, (left_idx, right_idx, score) in enumerate(matched_pairs):
            progress.Update(int((i + 1) / total_pairs * 100) if total_pairs > 0 else 100,
                          f"Computing diff {i + 1}/{total_pairs}...")
            wx.GetApp().Yield()

            # Get images for comparison
            left_page = self.left_doc.pages[left_idx]
            right_page = self.right_doc.pages[right_idx]

            if left_page.thumbnail and right_page.thumbnail:
                # Get exclusion zones for each side
                left_zones = self.exclusion_zones.get_zones_for("left")
                right_zones = self.exclusion_zones.get_zones_for("right")

                # Compare the images (with exclusion zones)
                diff_result = comparator.compare(
                    left_page.thumbnail, right_page.thumbnail,
                    exclusion_zones=left_zones
                )

                print(f"[DEBUG] Page {left_idx}: diff_score={diff_result.diff_score:.4f}, regions={len(diff_result.regions)}")

                # Apply highlighted images to thumbnails
                if diff_result.highlight_image:
                    # Convert PIL Image to wx.Bitmap
                    left_highlight_bitmap = pil_to_wxbitmap(diff_result.highlight_image)
                    self.left_panel.set_diff_result(left_idx, diff_result, left_highlight_bitmap)
                    print(f"[DEBUG] Left page {left_idx}: highlight applied, bitmap size={left_highlight_bitmap.GetWidth()}x{left_highlight_bitmap.GetHeight()}")

                    # Also create highlight for right side (compare right to left)
                    diff_result_right = comparator.compare(
                        right_page.thumbnail, left_page.thumbnail,
                        exclusion_zones=right_zones
                    )
                    if diff_result_right.highlight_image:
                        right_highlight_bitmap = pil_to_wxbitmap(diff_result_right.highlight_image)
                        self.right_panel.set_diff_result(right_idx, diff_result_right, right_highlight_bitmap)
                        print(f"[DEBUG] Right page {right_idx}: highlight applied")

        progress.Destroy()

        self.session.matching_result = self.matching_result

        # Update UI
        self.left_panel.update_match_status(self.matching_result)
        self.right_panel.update_match_status(self.matching_result)
        self._update_status()
        self._update_links()

    def _on_page_clicked(self, event: wx.CommandEvent) -> None:
        """Handle page click for manual linking."""
        index = event.GetInt()
        side = event.GetString()

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
        """Handle sync scrollbar movement."""
        if not self.left_doc or not self.right_doc:
            return

        value = self.sync_scrollbar.GetValue()
        percent = value / 100.0

        # Scroll both panels to the same percentage
        self._scroll_panel_to_percent(self.left_panel, percent)
        self._scroll_panel_to_percent(self.right_panel, percent)

        wx.CallAfter(self._update_links)

    def _on_sync_toggle(self, event: wx.CommandEvent) -> None:
        """Handle sync checkbox toggle."""
        if self.sync_checkbox.GetValue():
            # Sync both panels when enabled
            wx.CallAfter(self._update_sync_scrollbar)

    def _sync_scroll_from(self, source: DocumentPanel, target: DocumentPanel) -> None:
        """Sync scroll position from source to target panel."""
        # Get source scroll position as percentage
        source_range = source.GetScrollRange(wx.VERTICAL)
        source_pos = source.GetScrollPos(wx.VERTICAL)

        if source_range > 0:
            percent = source_pos / source_range
            target_range = target.GetScrollRange(wx.VERTICAL)
            target_pos = int(percent * target_range)
            target.Scroll(-1, target_pos)

    def _scroll_panel_to_percent(self, panel: DocumentPanel, percent: float) -> None:
        """Scroll panel to a percentage position."""
        scroll_range = panel.GetScrollRange(wx.VERTICAL)
        if scroll_range > 0:
            pos = int(percent * scroll_range)
            panel.Scroll(-1, pos)

    def _update_sync_scrollbar(self) -> None:
        """Update sync scrollbar position based on panel scroll."""
        if not hasattr(self, 'sync_scrollbar'):
            return

        # Use left panel position as reference
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
                links.append((left_pos, right_pos, match.status, match.similarity))

        self.link_overlay.set_links(links)

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
            self.frame._load_document(valid_files[1], "right")
        elif len(valid_files) == 1:
            if self.frame.left_doc is None:
                self.frame._load_document(valid_files[0], "left")
            elif self.frame.right_doc is None:
                self.frame._load_document(valid_files[0], "right")
            else:
                self.frame._load_document(valid_files[0], "left")

        return bool(valid_files)
