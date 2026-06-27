"""Настройка логирования и глобальная обработка исключений."""

from __future__ import annotations

import logging
import sys
import traceback
from tkinter import messagebox

from config import DEBUG_MODE, DRY_RUN, LOG_DIR


def setup_logging() -> None:
    log_file = LOG_DIR / "processor.log"
    handlers: list[logging.Handler] = [
        logging.FileHandler(log_file, encoding="utf-8"),
    ]
    if DEBUG_MODE:
        handlers.append(logging.StreamHandler())

    logging.basicConfig(
        level=logging.DEBUG if DEBUG_MODE else logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=handlers,
        force=True,
    )
    logging.info("=== Программа запущена ===")
    if DRY_RUN:
        logging.warning("РЕЖИМ DRY-RUN: файлы не будут сохранены")


def install_exception_handlers() -> None:
    def global_exception_handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, (SystemExit, KeyboardInterrupt)):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        error_msg = "".join(
            traceback.format_exception(exc_type, exc_value, exc_traceback)
        )
        logging.critical("Необработанное исключение:\n%s", error_msg)
        try:
            messagebox.showerror(
                "Критическая ошибка",
                "Программа столкнулась с ошибкой и будет закрыта.\n\n"
                f"Подробности записаны в лог-файл:\n{LOG_DIR / 'processor.log'}\n\n"
                f"Ошибка: {exc_type.__name__}: {exc_value}",
            )
        except Exception:
            pass
        sys.exit(1)

    sys.excepthook = global_exception_handler

    import threading

    def thread_exception_handler(args):
        global_exception_handler(args.exc_type, args.exc_value, args.exc_traceback)

    threading.excepthook = thread_exception_handler