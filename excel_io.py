"""Чтение Excel: движки, сканирование структуры, выборочная загрузка колонок."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import pandas as pd

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsb"}


def get_engine(filepath: Path) -> str | None:
    suffix = filepath.suffix.lower()
    if suffix == ".xlsx":
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    if suffix == ".xlsb":
        return "pyxlsb"
    return None


def list_sheets(filepath: Path) -> list[str]:
    engine = get_engine(filepath)
    xl = pd.ExcelFile(filepath, engine=engine)
    return xl.sheet_names


def list_columns(filepath: Path, sheet: int | str) -> list[str]:
    engine = get_engine(filepath)
    header_df = pd.read_excel(
        filepath, sheet_name=sheet, header=0, nrows=0, engine=engine
    )
    return header_df.columns.astype(str).tolist()


def scan_file(filepath: Path) -> dict[str, list[str]]:
    """Возвращает {имя_листа: [колонки]}."""
    engine = get_engine(filepath)
    xl = pd.ExcelFile(filepath, engine=engine)
    result: dict[str, list[str]] = {}
    for sheet in xl.sheet_names:
        header_df = pd.read_excel(
            filepath, sheet_name=sheet, header=0, nrows=0, engine=engine
        )
        result[sheet] = header_df.columns.astype(str).tolist()
    return result


def resolve_columns(
    all_columns: list[str],
    explicit: list[str],
    prefix: str,
    required: list[str] | None = None,
) -> list[str]:
    """Собирает список колонок: явные + по префиксу + обязательные."""
    selected: list[str] = []
    seen: set[str] = set()

    def add(col: str) -> None:
        if col in seen:
            return
        if col in all_columns:
            selected.append(col)
            seen.add(col)

    for col in required or []:
        add(col)

    if explicit:
        for col in explicit:
            add(col)
    elif prefix:
        prefix_lower = prefix.lower()
        for col in all_columns:
            if col.lower().startswith(prefix_lower):
                add(col)

    return selected


def read_selected_columns(
    filepath: Path,
    sheet: int | str,
    columns: list[str],
    dtype_dict: dict[str, Any] | None = None,
) -> pd.DataFrame:
    engine = get_engine(filepath)
    logging.debug(
        "Чтение %s, лист %s, колонки: %s", filepath.name, sheet, columns
    )
    return pd.read_excel(
        filepath,
        sheet_name=sheet,
        header=0,
        usecols=columns,
        dtype=dtype_dict,
        engine=engine,
    )


def sort_ai_columns(columns: list[str]) -> list[str]:
    def key(col: str) -> tuple[int, str]:
        lower = col.lower()
        if "ai" in lower:
            suffix = lower.split("ai")[-1]
            if suffix.isdigit():
                return (0, int(suffix))
        return (1, col)

    return sorted(columns, key=key)