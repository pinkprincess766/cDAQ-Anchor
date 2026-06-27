"""Тесты выбора колонок."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_io import resolve_columns, sort_ai_columns


def test_resolve_columns_by_prefix():
    all_cols = ["Time*", "pressure_ai0", "pressure_ai1", "Other"]
    selected = resolve_columns(
        all_cols,
        explicit=[],
        prefix="pressure_ai",
        required=["Time*"],
    )
    assert selected == ["Time*", "pressure_ai0", "pressure_ai1"]


def test_resolve_columns_explicit():
    all_cols = ["Time*", "pressure_ai0", "pressure_ai1"]
    selected = resolve_columns(
        all_cols,
        explicit=["pressure_ai1"],
        prefix="",
        required=["Time*"],
    )
    assert selected == ["Time*", "pressure_ai1"]


def test_sort_ai_columns():
    cols = ["pressure_ai10", "pressure_ai2", "pressure_ai1"]
    assert sort_ai_columns(cols) == [
        "pressure_ai1",
        "pressure_ai2",
        "pressure_ai10",
    ]