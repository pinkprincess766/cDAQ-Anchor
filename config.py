"""Конфигурация приложения: пути, профили обработки, настройки."""

from __future__ import annotations

import json
import logging
import os
import sys
from copy import deepcopy
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal

APP_NAME = "ExcelProcessor"

DEBUG_MODE = "--debug" in sys.argv
DRY_RUN = "--dry-run" in sys.argv

CriterionMode = Literal["closest", "above", "below"]
OutputFormat = Literal["xlsx", "csv", "both"]


def get_config_dir() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")
    else:
        base = Path.home() / ".config"
    path = Path(base) / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_log_dir() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
        path = Path(base) / APP_NAME / "Logs"
    else:
        path = Path.home() / ".local" / "share" / APP_NAME / "Logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


CONFIG_DIR = get_config_dir()
LOG_DIR = get_log_dir()
SETTINGS_FILE = CONFIG_DIR / "settings.json"


@dataclass
class ProcessingProfile:
    """Профиль обработки — определяет, что и откуда читать."""

    name: str = "Базовый профиль"
    p_suffix: str = "_p"
    t_suffix: str = "_t"
    p_sheet: int | str = 1
    t_sheet: int | str = 1
    p_time_column: str = "Time*"
    t_time_column: str = "Time*"
    p_criterion_column: str = "pressure_ai0"
    p_data_columns: list[str] = field(default_factory=list)
    t_data_columns: list[str] = field(default_factory=list)
    p_column_prefix: str = "pressure_ai"
    t_column_prefix: str = "thermo_ai"
    criterion_mode: CriterionMode = "closest"
    time_tolerance: float = 1e-6
    time_column_output: str = "Время"
    include_time_in_output: bool = True
    max_workers: int = 4
    output_format: OutputFormat = "xlsx"
    csv_separator: str = ";"
    csv_encoding: str = "utf-8-sig"
    write_dry_run_report: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ProcessingProfile:
        known = {f.name for f in cls.__dataclass_fields__.values()}  # type: ignore[attr-defined]
        filtered = {k: v for k, v in data.items() if k in known}
        return cls(**filtered)


DEFAULT_PROFILE = ProcessingProfile()


def default_settings() -> dict[str, Any]:
    return {
        "last_folder": "",
        "last_target": "",
        "active_profile": DEFAULT_PROFILE.name,
        "profiles": [DEFAULT_PROFILE.to_dict()],
    }


def load_settings() -> dict[str, Any]:
    if not SETTINGS_FILE.exists():
        return default_settings()
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as exc:
        logging.error("Ошибка загрузки настроек: %s", exc)
        return default_settings()

    if "profiles" not in data:
        legacy = ProcessingProfile.from_dict(data)
        return {
            "last_folder": data.get("last_folder", ""),
            "last_target": data.get("last_target", ""),
            "active_profile": legacy.name,
            "profiles": [legacy.to_dict()],
        }
    return data


def save_settings(settings: dict[str, Any]) -> None:
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
    except Exception as exc:
        logging.error("Ошибка сохранения настроек: %s", exc)


def get_profiles(settings: dict[str, Any]) -> list[ProcessingProfile]:
    raw = settings.get("profiles", [DEFAULT_PROFILE.to_dict()])
    return [ProcessingProfile.from_dict(p) for p in raw]


def get_active_profile(settings: dict[str, Any]) -> ProcessingProfile:
    profiles = get_profiles(settings)
    active_name = settings.get("active_profile", DEFAULT_PROFILE.name)
    for profile in profiles:
        if profile.name == active_name:
            return profile
    return profiles[0] if profiles else deepcopy(DEFAULT_PROFILE)


def set_active_profile(settings: dict[str, Any], profile: ProcessingProfile) -> None:
    profiles = get_profiles(settings)
    updated = False
    new_profiles = []
    for existing in profiles:
        if existing.name == profile.name:
            new_profiles.append(profile)
            updated = True
        else:
            new_profiles.append(existing)
    if not updated:
        new_profiles.append(profile)
    settings["profiles"] = [p.to_dict() for p in new_profiles]
    settings["active_profile"] = profile.name


def delete_profile(settings: dict[str, Any], name: str) -> bool:
    profiles = get_profiles(settings)
    if len(profiles) <= 1:
        return False
    remaining = [p for p in profiles if p.name != name]
    if len(remaining) == len(profiles):
        return False
    settings["profiles"] = [p.to_dict() for p in remaining]
    if settings.get("active_profile") == name:
        settings["active_profile"] = remaining[0].name
    return True


def duplicate_profile(settings: dict[str, Any], source_name: str, new_name: str) -> ProcessingProfile | None:
    profiles = get_profiles(settings)
    for profile in profiles:
        if profile.name == source_name:
            copy = deepcopy(profile)
            copy.name = new_name
            settings["profiles"] = [p.to_dict() for p in profiles] + [copy.to_dict()]
            settings["active_profile"] = new_name
            return copy
    return None


def rename_profile(settings: dict[str, Any], old_name: str, new_name: str) -> bool:
    profiles = get_profiles(settings)
    if any(p.name == new_name for p in profiles):
        return False
    updated = False
    new_profiles: list[ProcessingProfile] = []
    for profile in profiles:
        if profile.name == old_name:
            profile.name = new_name
            updated = True
        new_profiles.append(profile)
    if not updated:
        return False
    settings["profiles"] = [p.to_dict() for p in new_profiles]
    if settings.get("active_profile") == old_name:
        settings["active_profile"] = new_name
    return True


def export_profile_to_file(profile: ProcessingProfile, path: Path) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(profile.to_dict(), f, indent=4, ensure_ascii=False)


def import_profile_from_file(path: Path) -> ProcessingProfile:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return ProcessingProfile.from_dict(data)