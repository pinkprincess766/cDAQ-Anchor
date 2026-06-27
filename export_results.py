"""Сохранение результатов и отчётов."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import ProcessingProfile


def next_available_path(folder: Path, stem: str, suffix: str) -> Path:
    """Return a path that will not overwrite an existing file."""
    candidate = folder / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = folder / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def output_stem_from_results(results: list[dict]) -> str:
    from matching import extract_datetime_from_basename

    base_names = [r["_base_name"] for r in results]
    try:
        dated = [
            (extract_datetime_from_basename(name), name) for name in base_names
        ]
        dated.sort(key=lambda x: x[0])
        return dated[0][1][:10]
    except ValueError:
        return base_names[0][:10]


def save_dataframe(
    df: pd.DataFrame,
    folder: Path,
    stem: str,
    profile: ProcessingProfile,
) -> list[Path]:
    saved: list[Path] = []
    fmt = profile.output_format

    if fmt in ("xlsx", "both"):
        xlsx_path = next_available_path(folder, stem, ".xlsx")
        df.to_excel(xlsx_path, index=False)
        saved.append(xlsx_path)

    if fmt in ("csv", "both"):
        csv_path = next_available_path(folder, stem, ".csv")
        df.to_csv(
            csv_path,
            index=False,
            sep=profile.csv_separator,
            encoding=profile.csv_encoding,
        )
        saved.append(csv_path)

    return saved


def write_dry_run_report(
    folder: Path,
    stem: str,
    df: pd.DataFrame,
    profile: ProcessingProfile,
    processed: int,
    skipped: int,
    errors: list[str],
    pairs_total: int,
    elapsed_sec: float,
    target_value: float,
) -> Path:
    report_path = next_available_path(folder, f"{stem}_dry_run_report", ".json")
    preview = df.head(5).to_dict(orient="records")
    xlsx_path = next_available_path(folder, stem, ".xlsx")
    csv_path = next_available_path(folder, stem, ".csv")
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "profile": profile.name,
        "target_value": target_value,
        "pairs_total": pairs_total,
        "processed": processed,
        "skipped": skipped,
        "elapsed_sec": round(elapsed_sec, 3),
        "would_create": {
            "xlsx": xlsx_path.name if profile.output_format in ("xlsx", "both") else None,
            "csv": csv_path.name if profile.output_format in ("csv", "both") else None,
        },
        "output_format": profile.output_format,
        "columns": list(df.columns),
        "rows": len(df),
        "preview_first_5": preview,
        "errors": errors,
    }
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False, default=str)
    return report_path
