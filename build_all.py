"""Build script for both PySide6 and wxPython versions."""

import subprocess
import sys
from pathlib import Path


def main():
    """Build both EXE versions."""
    project_root = Path(__file__).parent

    print("=" * 60)
    print("Building ALL versions of PPT/PDF Comparator")
    print("=" * 60)

    # Build PySide6 version
    print("\n[1/2] Building PySide6 version...\n")
    result1 = subprocess.run(
        [sys.executable, str(project_root / "build_pyside.py")],
        cwd=str(project_root)
    )

    # Build wxPython version
    print("\n[2/2] Building wxPython version...\n")
    result2 = subprocess.run(
        [sys.executable, str(project_root / "build_wx.py")],
        cwd=str(project_root)
    )

    # Summary
    print("\n" + "=" * 60)
    print("BUILD SUMMARY")
    print("=" * 60)

    pyside_exe = project_root / "dist" / "PPTPDFComparator_PySide.exe"
    wx_exe = project_root / "dist" / "PPTPDFComparator_wx.exe"

    if pyside_exe.exists():
        size_mb = pyside_exe.stat().st_size / (1024 * 1024)
        print(f"PySide6: SUCCESS - {size_mb:.1f} MB")
    else:
        print("PySide6: FAILED")

    if wx_exe.exists():
        size_mb = wx_exe.stat().st_size / (1024 * 1024)
        print(f"wxPython: SUCCESS - {size_mb:.1f} MB")
    else:
        print("wxPython: FAILED")

    print("=" * 60)

    if result1.returncode != 0 or result2.returncode != 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
