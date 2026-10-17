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
    extract_time_str,
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

def test_find_criterion_row_missing_column():
    df = pd.DataFrame({"val": [1.0, 2.0]})
    with pytest.raises(KeyError):
        find_criterion_row(df, "absent", 1.0)


def test_find_criterion_row_no_numeric_values():
    df = pd.DataFrame({"val": ["abc", None]})
    with pytest.raises(ValueError, match="не содержит числовых"):
        find_criterion_row(df, "val", 1.0)


def test_find_criterion_row_above_without_candidates():
    df = pd.DataFrame({"val": [1.0, 2.0]})
    with pytest.raises(ValueError, match=">="):
        find_criterion_row(df, "val", 5.0, "above")


def test_find_criterion_row_below_without_candidates():
    df = pd.DataFrame({"val": [3.0, 4.0]})
    with pytest.raises(ValueError, match="<="):
        find_criterion_row(df, "val", 1.0, "below")


def test_find_criterion_row_unknown_mode():
    df = pd.DataFrame({"val": [1.0]})
    with pytest.raises(ValueError, match="Неизвестный режим"):
        find_criterion_row(df, "val", 1.0, "sideways")


def test_find_criterion_row_skips_non_numeric_rows():
    df = pd.DataFrame({"val": ["x", "4,5", "y", "4,0"]})
    idx, _ = find_criterion_row(df, "val", 4.4)
    assert idx == 1


def test_find_row_by_time_missing_column_or_nan_target():
    df = pd.DataFrame({"Time*": [0.1, 0.2]})
    assert find_row_by_time(df, "absent", 0.1) is None
    assert find_row_by_time(df, "Time*", float("nan")) is None


def test_extract_time_str_from_six_digit_run():
    assert extract_time_str(Path("26_06_2026_153045.xlsx")) == "15_30_45"


def test_extract_time_str_without_time():
    with pytest.raises(ValueError):
        extract_time_str(Path("no_time_here.xlsx"))
