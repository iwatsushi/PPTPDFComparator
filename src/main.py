"""Main entry point for PPT/PDF Comparator.

This module allows selection between PySide6 and wxPython GUIs.
"""

import sys
import os
import argparse

# PyInstaller用にパスを設定
if getattr(sys, 'frozen', False):
    # EXE実行時
    application_path = os.path.dirname(sys.executable)
    # srcディレクトリをパスに追加
    src_path = os.path.join(application_path, '_internal', 'src')
    if os.path.exists(src_path):
        sys.path.insert(0, src_path)
else:
    # 通常のPython実行時
    application_path = os.path.dirname(os.path.abspath(__file__))


def main():
    """Run the application."""
    parser = argparse.ArgumentParser(description="PPT/PDF Document Comparator")
    parser.add_argument(
        "--gui", "-g",
        choices=["pyside", "wx", "auto"],
        default="auto",
        help="GUI framework to use (default: auto)"
    )
    parser.add_argument(
        "--left", "-l",
        help="Path to left document"
    )
    parser.add_argument(
        "--right", "-r",
        help="Path to right document"
    )

    args = parser.parse_args()

    # Determine which GUI to use
    gui = args.gui

    if gui == "auto":
        # Try PySide6 first, fall back to wxPython
        try:
            import PySide6
            gui = "pyside"
        except ImportError:
            try:
                import wx
                gui = "wx"
            except ImportError:
                print("Error: Neither PySide6 nor wxPython is installed.")
                print("Please install one of them:")
                print("  pip install PySide6")
                print("  pip install wxPython")
                sys.exit(1)

    if gui == "pyside":
        # 絶対インポートを試みる
        try:
            from src.main_pyside import main as run_pyside
        except ImportError:
            from main_pyside import main as run_pyside
        run_pyside()
    else:
        # 絶対インポートを試みる
        try:
            from src.main_wx import main as run_wx
        except ImportError:
            from main_wx import main as run_wx
        run_wx()


if __name__ == "__main__":
    main()
