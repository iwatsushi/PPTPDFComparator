"""Build script for wxPython version of PPT/PDF Comparator."""

import subprocess
import sys
from pathlib import Path


def main():
    """Build wxPython EXE using PyInstaller."""
    print("=" * 60)
    print("Building wxPython version of PPT/PDF Comparator")
    print("=" * 60)

    # Get project root
    project_root = Path(__file__).parent
    src_dir = project_root / "src"

    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", "PPTPDFComparator_wx",
        "--paths", str(src_dir),
        # Hidden imports for core functionality
        "--hidden-import", "cv2",
        "--hidden-import", "numpy",
        "--hidden-import", "scipy.optimize",
        "--hidden-import", "imagehash",
        "--hidden-import", "fitz",
        "--hidden-import", "pptx",
        "--hidden-import", "comtypes",
        "--hidden-import", "comtypes.client",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "reportlab",
        "--hidden-import", "reportlab.lib",
        "--hidden-import", "reportlab.platypus",
        # Hidden imports for wxPython
        "--hidden-import", "wx",
        "--hidden-import", "wx.adv",
        "--hidden-import", "wx.lib",
        "--hidden-import", "wx.lib.newevent",
        # Hidden imports for src modules
        "--hidden-import", "src",
        "--hidden-import", "src.core",
        "--hidden-import", "src.core.document",
        "--hidden-import", "src.core.page_matcher",
        "--hidden-import", "src.core.image_comparator",
        "--hidden-import", "src.core.exclusion_zone",
        "--hidden-import", "src.core.session",
        "--hidden-import", "src.core.export",
        "--hidden-import", "src.utils",
        "--hidden-import", "src.utils.image_utils",
        "--hidden-import", "src.gui_wx",
        "--hidden-import", "src.gui_wx.main_window",
        # Also add without src prefix for fallback imports
        "--hidden-import", "core",
        "--hidden-import", "core.document",
        "--hidden-import", "core.page_matcher",
        "--hidden-import", "core.image_comparator",
        "--hidden-import", "core.exclusion_zone",
        "--hidden-import", "core.session",
        "--hidden-import", "core.export",
        "--hidden-import", "utils",
        "--hidden-import", "utils.image_utils",
        "--hidden-import", "gui_wx",
        "--hidden-import", "gui_wx.main_window",
        # Exclude unnecessary packages to reduce size
        "--exclude-module", "PySide6",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PyQt6",
        "--exclude-module", "torch",
        "--exclude-module", "tensorflow",
        "--exclude-module", "pandas",
        "--exclude-module", "matplotlib",
        "--exclude-module", "tkinter",
        # Entry point
        str(src_dir / "main_wx.py"),
    ]

    print("\nRunning PyInstaller...")
    print(f"Command: {' '.join(cmd)}\n")

    result = subprocess.run(cmd, cwd=str(project_root))

    if result.returncode == 0:
        exe_path = project_root / "dist" / "PPTPDFComparator_wx.exe"
        print("\n" + "=" * 60)
        print("Build successful!")
        print(f"EXE location: {exe_path}")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"EXE size: {size_mb:.1f} MB")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("Build FAILED!")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()
