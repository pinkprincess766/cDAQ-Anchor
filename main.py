"""Точка входа приложения ExcelProcessor."""

from __future__ import annotations

import sys
import tkinter as tk

# Опциональные зависимости (для .xls / .xlsb и PyInstaller)
for _optional in ("xlrd", "pyxlsb", "pytz", "dateutil", "numexpr", "bottleneck"):
    try:
        __import__(_optional)
    except ImportError:
        pass

from config import resource_path
from gui import App
from logging_setup import install_exception_handlers


def main() -> None:
    install_exception_handlers()
    root = tk.Tk()

    try:
        icon_path = resource_path("icon.ico")
        if __import__("os").path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
    sys.exit(0)