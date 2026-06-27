"""Тесты управления профилями."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from config import (
    ProcessingProfile,
    default_settings,
    delete_profile,
    duplicate_profile,
    export_profile_to_file,
    get_active_profile,
    import_profile_from_file,
    rename_profile,
)


def test_duplicate_and_delete_profile():
    settings = default_settings()
    duplicate_profile(settings, "Базовый профиль", "Тестовый")
    assert len(settings["profiles"]) == 2
    assert get_active_profile(settings).name == "Тестовый"

    assert delete_profile(settings, "Тестовый")
    assert len(settings["profiles"]) == 1
    assert not delete_profile(settings, "Базовый профиль")


def test_rename_profile():
    settings = default_settings()
    assert rename_profile(settings, "Базовый профиль", "Профиль стенда")
    assert get_active_profile(settings).name == "Профиль стенда"


def test_import_export_roundtrip(tmp_path):
    profile = ProcessingProfile(name="Export Test", output_format="csv")
    path = tmp_path / "profile.json"
    export_profile_to_file(profile, path)
    loaded = import_profile_from_file(path)
    assert loaded.name == "Export Test"
    assert loaded.output_format == "csv"