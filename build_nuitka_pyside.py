#!/usr/bin/env python3
"""Build script for PySide6 version using Nuitka."""

import subprocess
import sys
from pathlib import Path


def main():
    project_dir = Path(__file__).parent
    src_dir = project_dir / "src"
    dist_dir = project_dir / "dist"
    main_file = src_dir / "main_pyside.py"

    print("=" * 60)
    print("Building PySide6 version with Nuitka")
    print("=" * 60)

    # Nuitka command
    cmd = [
        sys.executable, "-m", "nuitka",

        # Output mode
        "--onefile",
        "--output-dir=" + str(dist_dir),
        "--output-filename=PPTPDFComparator_PySide_Nuitka.exe",

        # Windows GUI mode (no console)
        "--windows-console-mode=disable",

        # Assume yes for downloads (C compiler, etc.)
        "--assume-yes-for-downloads",

        # Enable plugins
        "--enable-plugin=pyside6",

        # Include packages
        "--include-package=src",
        "--include-package=src.core",
        "--include-package=src.gui_pyside",
        "--include-package=src.utils",

        # Include data for packages that need it
        "--include-package-data=reportlab",
        "--include-package-data=pptx",

        # Follow imports
        "--follow-imports",

        # Exclude unnecessary packages
        "--nofollow-import-to=wx",
        "--nofollow-import-to=wxPython",
        "--nofollow-import-to=torch",
        "--nofollow-import-to=tensorflow",
        "--nofollow-import-to=pandas",
        "--nofollow-import-to=matplotlib",
        "--nofollow-import-to=tkinter",

        # Main file
        str(main_file),
    ]

    print("\nRunning Nuitka...")
    print(f"Command: {' '.join(cmd)}")
    print()

    try:
        result = subprocess.run(cmd, check=True)

        exe_path = dist_dir / "PPTPDFComparator_PySide_Nuitka.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print()
            print("=" * 60)
            print("Build successful!")
            print(f"EXE location: {exe_path}")
            print(f"EXE size: {size_mb:.1f} MB")
            print("=" * 60)
        else:
            print("Build may have succeeded but EXE not found at expected location.")
            print("Check the dist directory for output files.")

    except subprocess.CalledProcessError as e:
        print(f"Build failed with error code: {e.returncode}")
        sys.exit(1)
    except FileNotFoundError:
        print("Nuitka not found. Install it with: pip install nuitka")
        sys.exit(1)


if __name__ == "__main__":
    main()
