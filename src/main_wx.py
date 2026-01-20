"""Entry point for wxPython version of PPT/PDF Comparator."""

import wx

try:
    from src.gui_wx import MainWindow
except ImportError:
    from gui_wx import MainWindow


def main():
    """Run the wxPython application."""
    app = wx.App()

    # Set app name
    app.SetAppName("PPT/PDF Comparator")

    window = MainWindow()
    window.Show()

    app.MainLoop()


if __name__ == "__main__":
    main()
