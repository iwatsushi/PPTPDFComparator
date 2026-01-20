"""PySide6 GUI implementation."""

try:
    from src.gui_pyside.main_window import MainWindow
except ImportError:
    from gui_pyside.main_window import MainWindow

__all__ = ["MainWindow"]
