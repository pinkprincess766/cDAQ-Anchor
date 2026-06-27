"""Оркестрация обработки пар файлов."""

from __future__ import annotations

import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Callable

import pandas as pd

from config import DRY_RUN, ProcessingProfile
from export_results import (
    next_available_path,
    output_stem_from_results,
    save_dataframe,
    write_dry_run_report,
)
from excel_io import list_columns, read_selected_columns, resolve_columns, sort_ai_columns
from matching import (
    extract_datetime_from_basename,
    extract_time_str,
    find_criterion_row,
    find_row_by_time,
    parse_numeric_series,
)

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int, int, str], None]


@dataclass
class PairFiles:
    base: str
    p_path: Path
    t_path: Path


@dataclass
class ProcessingReport:
    success: bool
    output_path: str | None
    output_paths: list[str]
    processed: int
    skipped: int
    errors: list[str]
    warnings: list[str]
    elapsed_sec: float
    dry_run_report_path: str | None = None


def discover_pairs(
    folder: Path,
    p_suffix: str = "_p",
    t_suffix: str = "_t",
) -> list[PairFiles]:
    from excel_io import SUPPORTED_EXTENSIONS

    all_files = [
        f
        for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    pairs: dict[str, dict[str, Path]] = {}

    for filepath in all_files:
        stem = filepath.stem
        if stem.endswith(p_suffix):
            base = stem[: -len(p_suffix)]
            pairs.setdefault(base, {})["p"] = filepath
        elif stem.endswith(t_suffix):
            base = stem[: -len(t_suffix)]
            pairs.setdefault(base, {})["t"] = filepath

    result: list[PairFiles] = []
    for base, mapping in sorted(pairs.items()):
        if "p" in mapping and "t" in mapping:
            result.append(PairFiles(base, mapping["p"], mapping["t"]))
    return result


def _read_side(
    filepath: Path,
    sheet: int | str,
    time_column: str,
    data_columns: list[str],
    prefix: str,
    extra_required: list[str] | None = None,
) -> pd.DataFrame:
    all_cols = list_columns(filepath, sheet)
    required = list({*(extra_required or []), time_column})
    selected = resolve_columns(all_cols, data_columns, prefix, required=required)
    if not selected:
        raise ValueError(f"Не выбрано ни одной колонки в {filepath.name}")
    return read_selected_columns(filepath, sheet, selected)


def process_pair(
    pair: PairFiles,
    target_value: float,
    profile: ProcessingProfile,
) -> dict | None:
    df_p = _read_side(
        pair.p_path,
        profile.p_sheet,
        profile.p_time_column,
        profile.p_data_columns,
        profile.p_column_prefix,
        extra_required=[profile.p_criterion_column],
    )

    if profile.p_criterion_column not in df_p.columns:
        raise KeyError(
            f"Нет колонки критерия '{profile.p_criterion_column}' в {pair.p_path.name}"
        )
    if profile.p_time_column not in df_p.columns:
        raise KeyError(
            f"Нет колонки времени '{profile.p_time_column}' в {pair.p_path.name}"
        )

    row_idx, row_p = find_criterion_row(
        df_p,
        profile.p_criterion_column,
        target_value,
        profile.criterion_mode,
    )
    time_values = parse_numeric_series(df_p[profile.p_time_column])
    time_val = float(time_values[row_idx])

    df_t = _read_side(
        pair.t_path,
        profile.t_sheet,
        profile.t_time_column,
        profile.t_data_columns,
        profile.t_column_prefix,
    )
    if profile.t_time_column not in df_t.columns:
        raise KeyError(
            f"Нет колонки времени '{profile.t_time_column}' в {pair.t_path.name}"
        )

    row_t = find_row_by_time(
        df_t,
        profile.t_time_column,
        time_val,
        profile.time_tolerance,
    )
    if row_t is None:
        raise ValueError(
            f"В {pair.t_path.name} время {time_val} не найдено "
            f"(допуск {profile.time_tolerance})"
        )

    record: dict = {}
    if profile.include_time_in_output:
        record[profile.p_time_column] = time_val

    t_data_cols = [
        c
        for c in df_t.columns
        if c != profile.t_time_column
        and (
            c in profile.t_data_columns
            or (
                not profile.t_data_columns
                and c.lower().startswith(profile.t_column_prefix.lower())
            )
        )
    ]
    for col in sort_ai_columns(t_data_cols):
        record[col] = row_t[col]

    p_data_cols = [
        c
        for c in df_p.columns
        if c != profile.p_time_column
        and c != profile.p_criterion_column
        and (
            c in profile.p_data_columns
            or (
                not profile.p_data_columns
                and c.lower().startswith(profile.p_column_prefix.lower())
            )
        )
    ]
    for col in sort_ai_columns(p_data_cols):
        record[col] = row_p[col]

    record[profile.p_criterion_column] = row_p[profile.p_criterion_column]
    record[profile.time_column_output] = extract_time_str(pair.p_path)
    record["_base_name"] = pair.base

    return record


def validate_setup(
    folder: Path,
    profile: ProcessingProfile,
    target_value: float | None = None,
) -> tuple[list[str], list[str]]:
    """Предварительная проверка. Возвращает (errors, warnings)."""
    errors: list[str] = []
    warnings: list[str] = []

    if not folder.exists():
        errors.append("Папка не найдена.")
        return errors, warnings
    if not folder.is_dir():
        errors.append("Указанный путь не является папкой.")
        return errors, warnings

    pairs = discover_pairs(folder, profile.p_suffix, profile.t_suffix)
    if not pairs:
        errors.append(
            f"Нет пар файлов с суффиксами «{profile.p_suffix}» / «{profile.t_suffix}»."
        )
        return errors, warnings

    if target_value is None:
        warnings.append("Целевое значение не задано — проверка значения критерия пропущена.")

    def add_missing(
        cols: list[str],
        required_cols: list[str],
        filename: str,
    ) -> None:
        for col in required_cols:
            if col not in cols:
                errors.append(f"Колонка «{col}» не найдена в {filename}.")

    for pair in pairs:
        try:
            cols_p = list_columns(pair.p_path, profile.p_sheet)
            p_required = [
                profile.p_criterion_column,
                profile.p_time_column,
                *profile.p_data_columns,
            ]
            add_missing(cols_p, p_required, pair.p_path.name)

            cols_t = list_columns(pair.t_path, profile.t_sheet)
            t_required = [profile.t_time_column, *profile.t_data_columns]
            add_missing(cols_t, t_required, pair.t_path.name)
        except Exception as exc:
            errors.append(f"Ошибка чтения пары {pair.base}: {exc}")

    orphan_p = orphan_t = 0
    mapping: dict[str, dict[str, Path]] = {}
    from excel_io import SUPPORTED_EXTENSIONS

    for filepath in folder.iterdir():
        if not filepath.is_file() or filepath.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        stem = filepath.stem
        if stem.endswith(profile.p_suffix):
            base = stem[: -len(profile.p_suffix)]
            mapping.setdefault(base, {})["p"] = filepath
        elif stem.endswith(profile.t_suffix):
            base = stem[: -len(profile.t_suffix)]
            mapping.setdefault(base, {})["t"] = filepath

    for base, files in mapping.items():
        if "p" in files and "t" not in files:
            orphan_p += 1
        elif "t" in files and "p" not in files:
            orphan_t += 1

    if orphan_p:
        warnings.append(f"Файлов _p без пары: {orphan_p}")
    if orphan_t:
        warnings.append(f"Файлов _t без пары: {orphan_t}")

    return errors, warnings


def _build_output_df(results: list[dict], profile: ProcessingProfile) -> pd.DataFrame:
    cleaned = []
    for rec in results:
        item = {k: v for k, v in rec.items() if not k.startswith("_")}
        cleaned.append(item)

    sample = cleaned[0]
    first_cols = []
    if profile.include_time_in_output and profile.p_time_column in sample:
        first_cols.append(profile.p_time_column)

    last_col = profile.time_column_output
    middle = [
        c
        for c in sample
        if c not in first_cols and c != last_col
    ]
    column_order = first_cols + middle + [last_col]
    return pd.DataFrame(cleaned)[column_order]


def run_processing(
    folder_path: str,
    target_value: float,
    profile: ProcessingProfile,
    log_callback: LogCallback,
    done_callback: Callable[[ProcessingReport], None],
    progress_callback: ProgressCallback | None = None,
    cancel_event: Event | None = None,
) -> None:
    total_start = time.time()
    errors: list[str] = []
    skipped = 0

    try:
        folder = Path(folder_path)
        if not folder.exists():
            log_callback("Ошибка: папка не найдена.")
            done_callback(
                ProcessingReport(
                    False, None, [], 0, 0, ["Папка не найдена"], [], 0
                )
            )
            return
        if not folder.is_dir():
            log_callback("Ошибка: указанный путь не является папкой.")
            done_callback(
                ProcessingReport(
                    False,
                    None,
                    [],
                    0,
                    0,
                    ["Указанный путь не является папкой"],
                    [],
                    0,
                )
            )
            return

        pairs = discover_pairs(folder, profile.p_suffix, profile.t_suffix)
        total = len(pairs)
        log_callback(f"Найдено полных пар: {total}")
        if total == 0:
            log_callback("Нет пар для обработки.")
            done_callback(
                ProcessingReport(False, None, [], 0, 0, ["Нет пар"], [], 0)
            )
            return

        results: list[dict] = []
        workers = max(1, min(profile.max_workers, total))

        def handle_pair(index: int, pair: PairFiles) -> tuple[int, dict | None, str | None]:
            if cancel_event and cancel_event.is_set():
                return index, None, "Отменено"
            try:
                rec = process_pair(pair, target_value, profile)
                if rec is None:
                    return index, None, "Пустой результат"
                try:
                    rec["_sort_key"] = extract_datetime_from_basename(rec["_base_name"])
                except ValueError:
                    rec["_sort_key"] = datetime.max
                return index, rec, None
            except Exception as exc:
                return index, None, str(exc)

        if workers == 1:
            for idx, pair in enumerate(pairs, start=1):
                if cancel_event and cancel_event.is_set():
                    log_callback("Обработка отменена пользователем.")
                    break
                log_callback(f"Обработка ({idx}/{total}): {pair.base}")
                if progress_callback:
                    progress_callback(idx - 1, total, pair.base)
                _, rec, err = handle_pair(idx, pair)
                if rec:
                    results.append(rec)
                else:
                    skipped += 1
                    msg = f"Пропуск {pair.base}: {err}"
                    errors.append(msg)
                    log_callback(msg)
                if progress_callback:
                    progress_callback(idx, total, pair.base)
        else:
            log_callback(f"Параллельная обработка ({workers} потоков)...")
            indexed_results: dict[int, dict] = {}
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {
                    executor.submit(handle_pair, idx, pair): (idx, pair)
                    for idx, pair in enumerate(pairs, start=1)
                }
                done_count = 0
                for future in as_completed(futures):
                    if cancel_event and cancel_event.is_set():
                        executor.shutdown(wait=False, cancel_futures=True)
                        log_callback("Обработка отменена пользователем.")
                        break
                    idx, pair = futures[future]
                    _, rec, err = future.result()
                    done_count += 1
                    if progress_callback:
                        progress_callback(done_count, total, pair.base)
                    if rec:
                        indexed_results[idx] = rec
                        log_callback(f"Готово ({done_count}/{total}): {pair.base}")
                    else:
                        skipped += 1
                        msg = f"Пропуск {pair.base}: {err}"
                        errors.append(msg)
                        log_callback(msg)
            results = [indexed_results[i] for i in sorted(indexed_results)]

        if cancel_event and cancel_event.is_set():
            done_callback(
                ProcessingReport(
                    False,
                    None,
                    [],
                    len(results),
                    skipped,
                    errors + ["Отменено"],
                    [],
                    time.time() - total_start,
                )
            )
            return

        if not results:
            log_callback("Нет успешно обработанных пар.")
            done_callback(
                ProcessingReport(
                    False, None, [], 0, skipped, errors, [], time.time() - total_start
                )
            )
            return

        results.sort(key=lambda x: x["_sort_key"])
        for rec in results:
            del rec["_sort_key"]

        stem = output_stem_from_results(results)
        output_df = _build_output_df(results, profile)
        saved_paths: list[Path] = []
        dry_run_report_path: Path | None = None

        if DRY_RUN:
            log_callback(f"[DRY-RUN] Базовое имя: {stem}")
            log_callback(
                f"[DRY-RUN] Формат: {profile.output_format} | "
                f"строк: {len(output_df)}, колонок: {len(output_df.columns)}"
            )
            if profile.output_format in ("xlsx", "both"):
                log_callback(f"[DRY-RUN] → {next_available_path(folder, stem, '.xlsx')}")
            if profile.output_format in ("csv", "both"):
                log_callback(f"[DRY-RUN] → {next_available_path(folder, stem, '.csv')}")
            if errors:
                log_callback(f"[DRY-RUN] Пропущено пар: {skipped}")
            if profile.write_dry_run_report:
                dry_run_report_path = write_dry_run_report(
                    folder,
                    stem,
                    output_df,
                    profile,
                    len(results),
                    skipped,
                    errors,
                    total,
                    time.time() - total_start,
                    target_value,
                )
                log_callback(f"[DRY-RUN] Отчёт: {dry_run_report_path}")
        else:
            saved_paths = save_dataframe(output_df, folder, stem, profile)
            for path in saved_paths:
                log_callback(f"Сохранено: {path}")

        elapsed = time.time() - total_start
        log_callback(
            f"Обработано: {len(results)}, пропущено: {skipped}, "
            f"время: {elapsed:.2f} сек"
        )
        logging.info(
            "Обработка завершена: %s пар за %.2f сек", len(results), elapsed
        )
        path_strings = [str(p) for p in saved_paths]
        done_callback(
            ProcessingReport(
                True,
                path_strings[0] if path_strings else None,
                path_strings,
                len(results),
                skipped,
                errors,
                [],
                elapsed,
                str(dry_run_report_path) if dry_run_report_path else None,
            )
        )
    except Exception as exc:
        log_callback(f"Критическая ошибка: {exc}")
        logging.error("Ошибка в run_processing", exc_info=True)
        done_callback(
            ProcessingReport(
                False,
                None,
                [],
                0,
                skipped,
                errors + [str(exc)],
                [],
                time.time() - total_start,
            )
        )
