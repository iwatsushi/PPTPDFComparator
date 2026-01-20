"""Image conversion utilities for GUI frameworks."""

from __future__ import annotations

from typing import TYPE_CHECKING, Tuple
from PIL import Image
import io

if TYPE_CHECKING:
    pass


def resize_to_fit(
    image: Image.Image,
    max_width: int,
    max_height: int,
    maintain_aspect: bool = True
) -> Image.Image:
    """Resize image to fit within bounds.

    Args:
        image: PIL Image to resize
        max_width: Maximum width
        max_height: Maximum height
        maintain_aspect: If True, maintain aspect ratio

    Returns:
        Resized PIL Image
    """
    if not maintain_aspect:
        return image.resize((max_width, max_height), Image.Resampling.LANCZOS)

    # Calculate scale to fit
    width_ratio = max_width / image.width
    height_ratio = max_height / image.height
    ratio = min(width_ratio, height_ratio)

    if ratio >= 1.0:
        return image  # Already fits

    new_width = int(image.width * ratio)
    new_height = int(image.height * ratio)

    return image.resize((new_width, new_height), Image.Resampling.LANCZOS)


def pil_to_qpixmap(image: Image.Image):
    """Convert PIL Image to QPixmap.

    Args:
        image: PIL Image

    Returns:
        QPixmap
    """
    from PySide6.QtGui import QImage, QPixmap

    # Convert to RGB if necessary
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # Get image data
    data = image.tobytes('raw', 'RGB')
    qimage = QImage(
        data,
        image.width,
        image.height,
        image.width * 3,
        QImage.Format.Format_RGB888
    )

    return QPixmap.fromImage(qimage)


def pil_to_wxbitmap(image: Image.Image):
    """Convert PIL Image to wx.Bitmap.

    Args:
        image: PIL Image

    Returns:
        wx.Bitmap
    """
    import wx

    # Convert to RGB if necessary
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # Create wx.Image
    wx_image = wx.Image(image.width, image.height)
    wx_image.SetData(image.tobytes())

    return wx_image.ConvertToBitmap()


def qpixmap_to_pil(pixmap) -> Image.Image:
    """Convert QPixmap to PIL Image.

    Args:
        pixmap: QPixmap

    Returns:
        PIL Image
    """
    from PySide6.QtCore import QBuffer, QIODevice

    buffer = QBuffer()
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    pixmap.save(buffer, "PNG")
    buffer.close()

    return Image.open(io.BytesIO(buffer.data().data()))


def wxbitmap_to_pil(bitmap) -> Image.Image:
    """Convert wx.Bitmap to PIL Image.

    Args:
        bitmap: wx.Bitmap

    Returns:
        PIL Image
    """
    import wx

    # Convert to wx.Image
    wx_image = bitmap.ConvertToImage()

    # Get dimensions
    width = wx_image.GetWidth()
    height = wx_image.GetHeight()

    # Get data
    data = wx_image.GetData()

    # Create PIL Image
    return Image.frombytes('RGB', (width, height), bytes(data))
