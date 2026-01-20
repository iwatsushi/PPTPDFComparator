"""Entry point for PySide6 version of PPT/PDF Comparator."""

import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

try:
    from src.gui_pyside import MainWindow
except ImportError:
    from gui_pyside import MainWindow


def main():
    """Run the PySide6 application."""
    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("PPT/PDF Comparator")
    app.setOrganizationName("PPTPDFComparator")

    # Set application style
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
