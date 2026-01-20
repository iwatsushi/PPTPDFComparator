"""wxPython GUI implementation."""

try:
    from src.gui_wx.main_window import MainWindow
except ImportError:
    from gui_wx.main_window import MainWindow

__all__ = ["MainWindow"]
