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
│   │   ├── session.py          # Session save/load
│   │   └── export.py           # PDF/HTML report export
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
├── dist/                       # Built EXE files
│   ├── PPTPDFComparator_PySide.exe
│   └── PPTPDFComparator_wx.exe
├── build_pyside.py             # PySide6 build script
├── build_wx.py                 # wxPython build script
├── build_all.py                # Build both versions
├── create_test_files.py        # Generate test files
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

### Using Build Scripts (Recommended)
```bash
# Install PyInstaller
pip install pyinstaller

# Build PySide6 version only
python build_pyside.py
# Output: dist/PPTPDFComparator_PySide.exe (~168MB)

# Build wxPython version only
python build_wx.py
# Output: dist/PPTPDFComparator_wx.exe (~136MB)

# Build both versions
python build_all.py
```

### Manual Build (Alternative)
```bash
# PySide6 version
pyinstaller --onefile --windowed --name PPTPDFComparator_PySide \
  --paths src \
  --hidden-import PySide6 --hidden-import cv2 --hidden-import numpy \
  --hidden-import scipy.optimize --hidden-import imagehash \
  --hidden-import fitz --hidden-import pptx --hidden-import comtypes.client \
  --hidden-import src.core --hidden-import src.utils --hidden-import src.gui_pyside \
  --exclude-module wx --exclude-module torch --exclude-module pandas \
  src/main_pyside.py

# wxPython version
pyinstaller --onefile --windowed --name PPTPDFComparator_wx \
  --paths src \
  --hidden-import wx --hidden-import cv2 --hidden-import numpy \
  --hidden-import scipy.optimize --hidden-import imagehash \
  --hidden-import fitz --hidden-import pptx --hidden-import comtypes.client \
  --hidden-import src.core --hidden-import src.utils --hidden-import src.gui_wx \
  --exclude-module PySide6 --exclude-module torch --exclude-module pandas \
  src/main_wx.py
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

## Performance Optimizations
- **PDF読み込み**: 並列レンダリング (ThreadPoolExecutor)
- **PowerPoint読み込み**: 一括エクスポート + 並列画像読み込み
- **キャッシュ**: ディスクキャッシュ (`~/.pptpdf_cache`) + 並列読み込み
- **比較処理**:
  - `compare_both()`: 1回の比較で両側のハイライト画像を生成 (2倍高速)
  - 並列比較処理 (ThreadPoolExecutor, 4-8ワーカー)
  - NumPyベクトル化によるハイライト生成最適化
- **ページマッチング**: pHashによる粗い比較 → SSIMによる詳細比較の2段階アプローチ

**期待される性能** (1000ページ比較):
- 読み込み: ~10秒
- 比較処理: ~10秒
- 合計: ~20秒

## Test Files
```bash
# Generate 1000-page test files
python create_test_files.py

# Creates:
# - A_1000.pdf, B_1000.pdf (1000 pages each)
# - A_1000.pptx, B_1000.pptx (1000 slides each)
# B versions have differences every 10 pages
```

## Notes
- 1000ページ規模のドキュメントを想定した効率化
- PowerPointファイルはMicrosoft PowerPointがインストールされている環境で最適動作
