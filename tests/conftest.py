"""Фикстуры для интеграционных тестов."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


@pytest.fixture
def sample_pair_folder(tmp_path: Path) -> Path:
    """Папка с парой _p/_t, данные на втором листе (index 1)."""
    p_sheet0 = pd.DataFrame({"meta": ["skip"]})
    p_sheet1 = pd.DataFrame(
        {
            "Time*": [0.1, 0.2, 0.3],
            "pressure_ai0": [4.0, 4.5, 5.0],
            "pressure_ai1": [10.0, 20.0, 30.0],
        }
    )
    t_sheet0 = pd.DataFrame({"meta": ["skip"]})
    t_sheet1 = pd.DataFrame(
        {
            "Time*": [0.1, 0.2, 0.3],
            "thermo_ai0": [100.0, 200.0, 300.0],
            "thermo_ai1": [101.0, 201.0, 301.0],
        }
    )

    p_path = tmp_path / "26_06_2026_100000_p.xlsx"
    t_path = tmp_path / "26_06_2026_100000_t.xlsx"

    with pd.ExcelWriter(p_path, engine="openpyxl") as writer:
        p_sheet0.to_excel(writer, sheet_name="Meta", index=False)
        p_sheet1.to_excel(writer, sheet_name="Data", index=False)

    with pd.ExcelWriter(t_path, engine="openpyxl") as writer:
        t_sheet0.to_excel(writer, sheet_name="Meta", index=False)
        t_sheet1.to_excel(writer, sheet_name="Data", index=False)

    return tmp_path