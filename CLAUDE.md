# PPT/PDF Comparator

## Overview
PowerPointとPDFドキュメントを視覚的に比較するGUIアプリケーション。
2つのドキュメントを左右に並べて表示し、ページ単位で画像比較を行い、差分をハイライト表示する。

## Tech Stack
- Language: Python 3.10+
- GUI Framework: PySide6 / wxPython (両方対応)
- Image Processing: OpenCV, Pillow, imagehash
- PDF Rendering: pdf2image (Poppler)
- PPT Rendering: python-pptx + LibreOffice

## Project Structure
```
PPTPDFComparator/
├── src/
│   ├── __init__.py
│   ├── main.py                 # Entry point (auto-select GUI)
│   ├── main_pyside.py          # PySide6 entry point
│   ├── main_wx.py              # wxPython entry point
│   │
│   ├── core/                   # Core logic (GUI-independent)
│   │   ├── document.py         # Document abstraction (PDF/PPT)
│   │   ├── page_matcher.py     # Page matching algorithm
│   │   ├── image_comparator.py # Image diff algorithm
│   │   ├── exclusion_zone.py   # Exclusion zone model
│   │   └── session.py          # Session save/load
│   │
│   ├── gui_pyside/             # PySide6 GUI
│   │   └── main_window.py
│   │
│   ├── gui_wx/                 # wxPython GUI
│   │   └── main_window.py
│   │
│   └── utils/
│       └── image_utils.py      # Image conversion utilities
│
├── tests/
├── pyproject.toml
└── CLAUDE.md
```

## Commands
```bash
# Install dependencies
pip install -e .

# Run with auto-detected GUI
python -m src.main

# Run with specific GUI
python -m src.main --gui pyside
python -m src.main --gui wx

# Run tests
pytest tests/
```

## Build EXE (Windows)
```bash
# Install PyInstaller
pip install pyinstaller

# Build standalone EXE
pyinstaller --onefile --windowed --name PPTPDFComparator \
  --paths src \
  --hidden-import wx --hidden-import cv2 --hidden-import numpy \
  --hidden-import scipy.optimize --hidden-import imagehash \
  --hidden-import fitz --hidden-import pptx --hidden-import comtypes.client \
  --hidden-import core --hidden-import utils --hidden-import gui_wx \
  --exclude-module PySide6 --exclude-module PyQt5 --exclude-module PyQt6 \
  --exclude-module torch --exclude-module pandas --exclude-module matplotlib \
  src/main.py

# Output: dist/PPTPDFComparator.exe
```

**Note:** PowerPoint files require Microsoft PowerPoint installed on the target machine.

## System Dependencies

### PDF Rendering
- **PyMuPDF (fitz)**: Included in Python dependencies, no external install needed

### PPTX Rendering (one of the following required)
- **Windows with Microsoft PowerPoint**: Uses COM automation (recommended)
- **LibreOffice**: Required if PowerPoint is not available
  - Windows: Download from https://www.libreoffice.org/download/
  - Mac: `brew install --cask libreoffice`
  - Linux: `apt install libreoffice`

**Note:** If neither PowerPoint nor LibreOffice is installed, PPTX files will show gray placeholder images.

## Key Features
- ドラッグ&ドロップでファイル読込
- ページマッチング (pHash + Hungarian algorithm)
- 除外エリア設定 (ページ番号等を無視)
- 手動リンク編集
- セッション保存/読込
- PDF/HTMLレポートエクスポート

## Coding Conventions
- Type hints: 全ての関数に型ヒントを記載
- Docstrings: Google style
- Naming: snake_case for functions/variables, PascalCase for classes

## Notes
- 1000ページ規模のドキュメントを想定した効率化
- pHashによる粗い比較 → SSIMによる詳細比較の2段階アプローチ
