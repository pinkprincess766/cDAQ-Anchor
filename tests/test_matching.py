"""Тесты сопоставления строк и парсинга имён."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from matching import (
    extract_datetime_from_basename,
    find_criterion_row,
    find_row_by_time,
    parse_numeric_series,
)


def test_parse_numeric_series_with_comma():
    series = pd.Series(["1,5", "2,0", "abc"])
    values = parse_numeric_series(series)
    assert np.isclose(values[0], 1.5)
    assert np.isclose(values[1], 2.0)
    assert np.isnan(values[2])


def test_find_criterion_row_closest():
    df = pd.DataFrame({"pressure_ai0": [1.0, 4.5, 5.0], "Time*": [0.1, 0.2, 0.3]})
    idx, row = find_criterion_row(df, "pressure_ai0", 4.47, "closest")
    assert idx == 1
    assert row["pressure_ai0"] == 4.5


def test_find_criterion_row_above():
    df = pd.DataFrame({"val": [1.0, 4.5, 5.0]})
    idx, row = find_criterion_row(df, "val", 4.47, "above")
    assert idx == 1
    assert row["val"] == 4.5


def test_find_criterion_row_below():
    df = pd.DataFrame({"val": [1.0, 4.5, 5.0]})
    idx, row = find_criterion_row(df, "val", 4.6, "below")
    assert idx == 1
    assert row["val"] == 4.5


def test_find_row_by_time_with_tolerance():
    df = pd.DataFrame({"Time*": [0.1, 0.2000000001, 0.3]})
    row = find_row_by_time(df, "Time*", 0.2, tolerance=1e-6)
    assert row is not None
    assert np.isclose(row["Time*"], 0.2000000001)


def test_find_row_by_time_no_match():
    df = pd.DataFrame({"Time*": [0.1, 0.5, 0.9]})
    assert find_row_by_time(df, "Time*", 0.2, tolerance=1e-6) is None


def test_extract_datetime_from_basename():
    dt = extract_datetime_from_basename("26_06_2026_153045")
    assert dt == datetime(2026, 6, 26, 15, 30, 45)


def test_extract_datetime_invalid():
    with pytest.raises(ValueError):
        extract_datetime_from_basename("bad_name")