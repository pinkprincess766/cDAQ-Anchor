"""Запуск приложения с автоматической подготовкой локального .venv."""

from __future__ import annotations

import os
import platform
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SETUP_ONLY = "--setup-only" in sys.argv


def _venv_python() -> Path:
    if platform.system() == "Windows":
        return ROOT / ".venv" / "Scripts" / "python.exe"
    return ROOT / ".venv" / "bin" / "python"


def _in_virtualenv() -> bool:
    return sys.prefix != getattr(sys, "base_prefix", sys.prefix)


def _create_venv(venv_python: Path) -> None:
    print("Локальное окружение .venv не найдено. Создаю...", flush=True)
    subprocess.check_call([sys.executable, "-m", "venv", str(ROOT / ".venv")])
    if not venv_python.exists():
        raise RuntimeError(f"Python внутри .venv не найден: {venv_python}")


def _install_dependencies(venv_python: Path) -> None:
    print("Устанавливаю зависимости...", flush=True)
    subprocess.check_call(
        [str(venv_python), "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")]
    )


def _reexec_with_venv() -> None:
    if _in_virtualenv():
        if SETUP_ONLY:
            print("Окружение готово.", flush=True)
            sys.exit(0)
        return
    venv_python = _venv_python()
    if not venv_python.exists():
        try:
            _create_venv(venv_python)
            _install_dependencies(venv_python)
        except Exception as exc:
            print(
                "Не удалось автоматически подготовить окружение.\n\n"
                f"Ошибка: {exc}\n\n"
                "Попробуйте вручную:\n"
                "  python3 -m venv .venv                 # macOS / Linux\n"
                "  .venv/bin/python -m pip install -r requirements.txt\n\n"
                "  py -m venv .venv                      # Windows\n"
                "  .venv\\Scripts\\python -m pip install -r requirements.txt"
            )
            sys.exit(1)
    if SETUP_ONLY:
        print("Окружение готово.", flush=True)
        sys.exit(0)
    args = [str(venv_python), str(ROOT / "main.py"), *sys.argv[1:]]
    os.execv(str(venv_python), args)


_reexec_with_venv()

try:
    import pandas  # noqa: F401
    import openpyxl  # noqa: F401
except ImportError:
    print(
        "Не установлены зависимости.\n\n"
        "Выполните в папке проекта:\n"
        "  .venv/bin/python -m pip install -r requirements.txt      # macOS / Linux\n"
        "  .venv\\Scripts\\python -m pip install -r requirements.txt # Windows\n\n"
        "Запуск:\n"
        "  python3 start.py       # macOS / Linux\n"
        "  py start.py            # Windows, если установлен Python Launcher\n"
        "  python start.py        # Windows, если python есть в PATH"
    )
    sys.exit(1)

from main import main  # noqa: E402

if __name__ == "__main__":
    main()
