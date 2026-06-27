"""Интеграционные тесты полного пайплайна."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from config import ProcessingProfile
from export_results import output_stem_from_results, save_dataframe
from processing import discover_pairs, process_pair, validate_setup


@pytest.fixture
def data_profile() -> ProcessingProfile:
    return ProcessingProfile(
        p_sheet="Data",
        t_sheet="Data",
        max_workers=1,
        output_format="both",
    )


def test_discover_and_validate(sample_pair_folder, data_profile):
    pairs = discover_pairs(sample_pair_folder)
    assert len(pairs) == 1
    errors, warnings = validate_setup(sample_pair_folder, data_profile, 4.47)
    assert errors == []
    assert warnings == []


def test_process_pair_finds_closest_pressure(sample_pair_folder, data_profile):
    pairs = discover_pairs(sample_pair_folder)
    record = process_pair(pairs[0], 4.47, data_profile)
    assert record is not None
    assert record["pressure_ai0"] == 4.5
    assert record["Time*"] == pytest.approx(0.2)
    assert record["thermo_ai0"] == 200.0
    assert record["Время"] == "10_00_00"


def test_full_export_pipeline(sample_pair_folder, data_profile, tmp_path):
    pairs = discover_pairs(sample_pair_folder)
    record = process_pair(pairs[0], 4.47, data_profile)
    assert record
    record["_sort_key"] = 0
    del record["_sort_key"]

    stem = output_stem_from_results([record])
    assert stem == "26_06_2026"

    from processing import _build_output_df

    df = _build_output_df([record], data_profile)
    out_dir = tmp_path / "output"
    out_dir.mkdir()
    paths = save_dataframe(df, out_dir, stem, data_profile)
    assert len(paths) == 2

    xlsx_df = pd.read_excel(paths[0])
    csv_df = pd.read_csv(paths[1], sep=";")
    assert len(xlsx_df) == 1
    assert len(csv_df) == 1
    assert xlsx_df["pressure_ai0"].iloc[0] == 4.5


def test_validate_detects_missing_criterion_column(sample_pair_folder):
    bad_profile = ProcessingProfile(
        p_sheet="Data",
        t_sheet="Data",
        p_criterion_column="MISSING_COL",
    )
    errors, _ = validate_setup(sample_pair_folder, bad_profile, 4.47)
    assert any("MISSING_COL" in e for e in errors)


def test_validate_checks_all_pairs(sample_pair_folder, data_profile):
    p_path = sample_pair_folder / "27_06_2026_100000_p.xlsx"
    t_path = sample_pair_folder / "27_06_2026_100000_t.xlsx"

    with pd.ExcelWriter(p_path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "Time*": [0.1],
                "WRONG_PRESSURE_COLUMN": [4.5],
            }
        ).to_excel(writer, sheet_name="Data", index=False)

    with pd.ExcelWriter(t_path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "Time*": [0.1],
                "thermo_ai0": [200.0],
            }
        ).to_excel(writer, sheet_name="Data", index=False)

    errors, warnings = validate_setup(sample_pair_folder, data_profile, 4.47)
    assert warnings == []
    assert any("27_06_2026_100000_p.xlsx" in e for e in errors)


def test_process_pair_accepts_comma_decimal_time(tmp_path, data_profile):
    p_path = tmp_path / "26_06_2026_100000_p.xlsx"
    t_path = tmp_path / "26_06_2026_100000_t.xlsx"

    with pd.ExcelWriter(p_path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "Time*": ["0,1", "0,2"],
                "pressure_ai0": ["4,0", "4,5"],
            }
        ).to_excel(writer, sheet_name="Data", index=False)

    with pd.ExcelWriter(t_path, engine="openpyxl") as writer:
        pd.DataFrame(
            {
                "Time*": ["0,1", "0,2"],
                "thermo_ai0": [100.0, 200.0],
            }
        ).to_excel(writer, sheet_name="Data", index=False)

    pair = discover_pairs(tmp_path)[0]
    record = process_pair(pair, 4.47, data_profile)

    assert record["Time*"] == pytest.approx(0.2)
    assert record["thermo_ai0"] == 200.0


def test_dry_run_report_written(sample_pair_folder, data_profile, monkeypatch):
    monkeypatch.setattr("config.DRY_RUN", True, raising=False)
    monkeypatch.setattr("processing.DRY_RUN", True, raising=False)

    from export_results import write_dry_run_report
    from processing import _build_output_df

    pairs = discover_pairs(sample_pair_folder)
    record = process_pair(pairs[0], 4.47, data_profile)
    df = _build_output_df([record], data_profile)
    stem = output_stem_from_results([record])

    report_path = write_dry_run_report(
        sample_pair_folder,
        stem,
        df,
        data_profile,
        processed=1,
        skipped=0,
        errors=[],
        pairs_total=1,
        elapsed_sec=0.5,
        target_value=4.47,
    )
    assert report_path.exists()
    content = report_path.read_text(encoding="utf-8")
    assert "dry_run_report" in report_path.name
    assert "4.47" in content
