"""Тесты обработки пар и формирования выходного файла."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from export_results import output_stem_from_results, save_dataframe
from config import ProcessingProfile
from processing import discover_pairs, validate_setup
import pandas as pd


def test_output_stem_uses_earliest_date():
    results = [
        {"_base_name": "27_06_2026_120000"},
        {"_base_name": "26_06_2026_153045"},
    ]
    assert output_stem_from_results(results) == "26_06_2026"


def test_output_stem_fallback():
    results = [{"_base_name": "invalid"}, {"_base_name": "also_bad"}]
    assert output_stem_from_results(results) == "invalid"


def test_save_dataframe_csv(tmp_path):
    df = pd.DataFrame({"a": [1], "b": [2]})
    profile = ProcessingProfile(output_format="csv", csv_separator=";")
    paths = save_dataframe(df, tmp_path, "test_out", profile)
    assert len(paths) == 1
    assert paths[0].suffix == ".csv"
    assert "1;2" in paths[0].read_text(encoding="utf-8-sig")


def test_save_dataframe_does_not_overwrite_existing_file(tmp_path):
    existing = tmp_path / "test_out.csv"
    existing.write_text("keep me", encoding="utf-8")
    df = pd.DataFrame({"a": [1]})
    profile = ProcessingProfile(output_format="csv", csv_separator=";")

    paths = save_dataframe(df, tmp_path, "test_out", profile)

    assert paths[0].name == "test_out_1.csv"
    assert existing.read_text(encoding="utf-8") == "keep me"


def test_save_dataframe_both(tmp_path):
    df = pd.DataFrame({"a": [1]})
    profile = ProcessingProfile(output_format="both")
    paths = save_dataframe(df, tmp_path, "both_out", profile)
    assert len(paths) == 2
    assert {p.suffix for p in paths} == {".xlsx", ".csv"}


def test_validate_setup_no_pairs(tmp_path):
    errors, warnings = validate_setup(tmp_path, ProcessingProfile(), 4.47)
    assert errors
    assert warnings == []


def test_discover_pairs_empty_dir(tmp_path: Path):
    assert discover_pairs(tmp_path) == []


def test_discover_pairs_finds_p_and_t(tmp_path: Path):
    (tmp_path / "26_06_2026_100000_p.xlsx").write_bytes(b"")
    (tmp_path / "26_06_2026_100000_t.xlsx").write_bytes(b"")
    (tmp_path / "orphan_p.xlsx").write_bytes(b"")

    pairs = discover_pairs(tmp_path)
    assert len(pairs) == 1
    assert pairs[0].base == "26_06_2026_100000"
