"""Сопоставление строк: критерий давления и синхронизация по времени."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from config import CriterionMode


def parse_numeric_series(series: pd.Series) -> np.ndarray:
    """Преобразует колонку в float, учитывая запятую как разделитель."""
    if pd.api.types.is_numeric_dtype(series):
        return series.to_numpy(dtype=float, copy=False)
    cleaned = series.astype(str).str.replace(",", ".", regex=False)
    return pd.to_numeric(cleaned, errors="coerce").to_numpy(dtype=float)


def find_criterion_row(
    df: pd.DataFrame,
    column: str,
    target_value: float,
    mode: CriterionMode = "closest",
) -> tuple[int, pd.Series]:
    if column not in df.columns:
        raise KeyError(f"Колонка критерия '{column}' не найдена")

    values = parse_numeric_series(df[column])
    valid_mask = ~np.isnan(values)
    if not valid_mask.any():
        raise ValueError(f"Колонка '{column}' не содержит числовых значений")

    valid_indices = np.flatnonzero(valid_mask)
    valid_values = values[valid_mask]

    if mode == "closest":
        diff = np.abs(valid_values - target_value)
        best_local = int(np.argmin(diff))
    elif mode == "above":
        above = valid_values >= target_value
        if not above.any():
            raise ValueError(
                f"Нет значений >= {target_value} в колонке '{column}'"
            )
        candidates = valid_values[above]
        best_local = int(np.argmin(candidates - target_value))
        valid_indices = valid_indices[above]
        valid_values = candidates
    elif mode == "below":
        below = valid_values <= target_value
        if not below.any():
            raise ValueError(
                f"Нет значений <= {target_value} в колонке '{column}'"
            )
        candidates = valid_values[below]
        best_local = int(np.argmin(target_value - candidates))
        valid_indices = valid_indices[below]
        valid_values = candidates
    else:
        raise ValueError(f"Неизвестный режим критерия: {mode}")

    row_idx = int(valid_indices[best_local])
    return row_idx, df.iloc[row_idx]


def find_row_by_time(
    df: pd.DataFrame,
    time_column: str,
    target_time: float,
    tolerance: float = 1e-6,
) -> pd.Series | None:
    if time_column not in df.columns:
        return None

    times = parse_numeric_series(df[time_column])
    if np.isnan(target_time):
        return None

    diff = np.abs(times - target_time)
    finite_mask = ~np.isnan(diff)
    if not finite_mask.any():
        return None

    best_idx = int(np.argmin(np.where(finite_mask, diff, np.inf)))
    if diff[best_idx] > tolerance:
        return None

    return df.iloc[best_idx]


def extract_datetime_from_basename(basename: str) -> datetime:
    parts = basename.split("_")
    if len(parts) == 4:
        day, month, year, time_str = parts
        if len(time_str) == 6 and time_str.isdigit():
            return datetime(
                int(year),
                int(month),
                int(day),
                int(time_str[:2]),
                int(time_str[2:4]),
                int(time_str[4:6]),
            )
    raise ValueError(f"Не удалось разобрать имя: {basename}")


def extract_time_str(filepath: Path) -> str:
    stem = filepath.stem
    match = re.search(r"(\d{6})", stem)
    if match:
        time_str = match.group(1)
        return f"{time_str[:2]}_{time_str[2:4]}_{time_str[4:6]}"

    base = stem[:-2] if len(stem) > 2 else stem
    time_part = base.split("_")[-1]
    if len(time_part) == 6 and time_part.isdigit():
        return f"{time_part[:2]}_{time_part[2:4]}_{time_part[4:6]}"

    raise ValueError(f"Не могу извлечь время из {stem}")
