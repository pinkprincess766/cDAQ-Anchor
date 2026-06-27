"""Графический интерфейс обработчика пар Excel."""

from __future__ import annotations

import logging
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk

from config import (
    DEBUG_MODE,
    DRY_RUN,
    ProcessingProfile,
    delete_profile,
    duplicate_profile,
    export_profile_to_file,
    get_active_profile,
    get_profiles,
    import_profile_from_file,
    load_settings,
    rename_profile,
    save_settings,
    set_active_profile,
)
from excel_io import SUPPORTED_EXTENSIONS, scan_file
from logging_setup import setup_logging
from processing import ProcessingReport, discover_pairs, run_processing, validate_setup


class ColumnSelector(ttk.Frame):
    """Мультивыбор колонок: список + кнопки добавить/убрать."""

    def __init__(self, master, label: str, **kwargs):
        super().__init__(master, **kwargs)
        self._all_columns: list[str] = []
        ttk.Label(self, text=label).pack(anchor=tk.W)
        row = ttk.Frame(self)
        row.pack(fill=tk.BOTH, expand=True)

        self.available = tk.Listbox(row, selectmode=tk.EXTENDED, height=6, exportselection=False)
        self.selected = tk.Listbox(row, selectmode=tk.EXTENDED, height=6, exportselection=False)
        self.available.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.selected.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(4, 0))

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, pady=4)
        ttk.Button(btns, text="→", width=4, command=self._add).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="←", width=4, command=self._remove).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="По префиксу", command=self._add_by_prefix).pack(side=tk.LEFT, padx=8)

        self._prefix_provider: callable = lambda: ""

    def set_prefix_provider(self, provider: callable) -> None:
        self._prefix_provider = provider

    def set_available(self, columns: list[str]) -> None:
        self._all_columns = list(columns)
        self._refresh_available()

    def set_selected(self, columns: list[str]) -> None:
        self.selected.delete(0, tk.END)
        for col in columns:
            self.selected.insert(tk.END, col)
        self._refresh_available()

    def get_selected(self) -> list[str]:
        return [self.selected.get(i) for i in range(self.selected.size())]

    def _add(self) -> None:
        indices = list(self.available.curselection())
        for idx in indices:
            col = self.available.get(idx)
            if col not in self.get_selected():
                self.selected.insert(tk.END, col)
        self._refresh_available()

    def _remove(self) -> None:
        indices = list(self.selected.curselection())
        for idx in reversed(indices):
            self.selected.delete(idx)
        self._refresh_available()

    def _add_by_prefix(self) -> None:
        prefix = self._prefix_provider().strip().lower()
        if not prefix:
            return
        for i in range(self.available.size()):
            col = self.available.get(i)
            if col.lower().startswith(prefix) and col not in self.get_selected():
                self.selected.insert(tk.END, col)
        self._refresh_available()

    def _refresh_available(self) -> None:
        selected_set = set(self.get_selected())
        self.available.delete(0, tk.END)
        for col in self._all_columns:
            if col not in selected_set:
                self.available.insert(tk.END, col)


class App:
    def __init__(self, root: tk.Tk):
        setup_logging()
        self.root = root
        self.settings = load_settings()
        self.profile = get_active_profile(self.settings)
        self.processing_thread: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self._scan_data: dict[str, dict[str, list[str]]] = {}

        title = "Excel Pair Processor"
        if DRY_RUN:
            title += " [DRY-RUN]"
        if DEBUG_MODE:
            title += " [DEBUG]"
        root.title(title)
        root.geometry("900x780")
        root.minsize(800, 650)

        self._build_ui()
        self._load_profile_to_ui(self.profile)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab_main = ttk.Frame(notebook)
        tab_config = ttk.Frame(notebook)
        notebook.add(tab_main, text="Обработка")
        notebook.add(tab_config, text="Настройки")

        self._build_main_tab(tab_main)
        self._build_config_tab(tab_config)

        self.status_var = tk.StringVar(value="Готов")
        ttk.Label(
            self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W
        ).pack(fill=tk.X, side=tk.BOTTOM, padx=8, pady=4)

    def _build_main_tab(self, parent: ttk.Frame) -> None:
        folder_frame = ttk.LabelFrame(parent, text="1. Папка с файлами", padding=6)
        folder_frame.pack(fill=tk.X, pady=4)

        self.folder_path = tk.StringVar(value=self.settings.get("last_folder", ""))
        ttk.Entry(folder_frame, textvariable=self.folder_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6)
        )
        ttk.Button(folder_frame, text="Обзор", command=self.browse_folder).pack(side=tk.RIGHT)
        ttk.Button(folder_frame, text="Сканировать", command=self.scan_folder).pack(
            side=tk.RIGHT, padx=4
        )

        target_frame = ttk.LabelFrame(parent, text="2. Целевое значение критерия", padding=6)
        target_frame.pack(fill=tk.X, pady=4)

        self.target_value = tk.StringVar(value=self.settings.get("last_target", ""))
        ttk.Entry(target_frame, textvariable=self.target_value, width=16).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Label(target_frame, text="(пример: 4.47)").pack(side=tk.LEFT)

        profile_row = ttk.LabelFrame(parent, text="Профиль обработки", padding=6)
        profile_row.pack(fill=tk.X, pady=4)
        inner = ttk.Frame(profile_row)
        inner.pack(fill=tk.X)
        ttk.Label(inner, text="Активный:").pack(side=tk.LEFT, padx=4)
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(
            inner, textvariable=self.profile_var, state="readonly", width=32
        )
        self.profile_combo.pack(side=tk.LEFT, padx=4)
        self.profile_combo.bind("<<ComboboxSelected>>", self.on_profile_changed)

        prof_btns = ttk.Frame(profile_row)
        prof_btns.pack(fill=tk.X, pady=4)
        ttk.Button(prof_btns, text="Новый", command=self.new_profile).pack(side=tk.LEFT, padx=2)
        ttk.Button(prof_btns, text="Копия", command=self.duplicate_profile_ui).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(prof_btns, text="Удалить", command=self.delete_profile_ui).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(prof_btns, text="Импорт", command=self.import_profile_ui).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(prof_btns, text="Экспорт", command=self.export_profile_ui).pack(
            side=tk.LEFT, padx=2
        )
        self._refresh_profile_list()

        export_row = ttk.Frame(parent)
        export_row.pack(fill=tk.X, pady=2)
        ttk.Label(export_row, text="Формат вывода:").pack(side=tk.LEFT, padx=4)
        self.output_format_var = tk.StringVar(value="xlsx")
        ttk.Combobox(
            export_row,
            textvariable=self.output_format_var,
            values=["xlsx", "csv", "both"],
            state="readonly",
            width=10,
        ).pack(side=tk.LEFT, padx=4)
        ttk.Label(export_row, text="CSV разделитель:").pack(side=tk.LEFT, padx=(12, 4))
        self.csv_sep_var = tk.StringVar(value=";")
        ttk.Entry(export_row, textvariable=self.csv_sep_var, width=4).pack(side=tk.LEFT)

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill=tk.X, pady=8)
        ttk.Button(btn_row, text="Проверить", command=self.validate_before_run).pack(
            side=tk.LEFT, padx=4
        )
        self.run_btn = ttk.Button(btn_row, text="ЗАПУСТИТЬ", command=self.start_processing)
        self.run_btn.pack(side=tk.LEFT, padx=4)
        self.cancel_btn = ttk.Button(
            btn_row, text="ОТМЕНА", command=self.cancel_processing, state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=4)

        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.pack(fill=tk.X, padx=4, pady=4)

        log_frame = ttk.LabelFrame(parent, text="3. Ход выполнения", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=4)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=18)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def _build_config_tab(self, parent: ttk.Frame) -> None:
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor=tk.NW)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        general = ttk.LabelFrame(scroll_frame, text="Общие", padding=6)
        general.pack(fill=tk.X, padx=4, pady=4)

        self.profile_name_var = tk.StringVar()
        self._labeled_entry(general, "Имя профиля:", self.profile_name_var, 0)

        self.p_suffix_var = tk.StringVar()
        self.t_suffix_var = tk.StringVar()
        self.criterion_mode_var = tk.StringVar()
        self.time_tolerance_var = tk.StringVar()
        self.max_workers_var = tk.StringVar()

        self._labeled_entry(general, "Суффикс _p:", self.p_suffix_var, 1)
        self._labeled_entry(general, "Суффикс _t:", self.t_suffix_var, 2)
        ttk.Label(general, text="Режим критерия:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(
            general,
            textvariable=self.criterion_mode_var,
            values=["closest", "above", "below"],
            state="readonly",
            width=12,
        ).grid(row=3, column=1, sticky=tk.W, pady=2)
        self._labeled_entry(general, "Допуск Time*:", self.time_tolerance_var, 4)
        self._labeled_entry(general, "Потоков:", self.max_workers_var, 5)

        self.dry_run_report_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            general,
            text="Писать JSON-отчёт в dry-run",
            variable=self.dry_run_report_var,
        ).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=4)

        sheets = ttk.LabelFrame(scroll_frame, text="Листы", padding=6)
        sheets.pack(fill=tk.X, padx=4, pady=4)

        self.p_sheet_var = tk.StringVar()
        self.t_sheet_var = tk.StringVar()
        ttk.Label(sheets, text="Лист _p:").grid(row=0, column=0, sticky=tk.W)
        self.p_sheet_combo = ttk.Combobox(sheets, textvariable=self.p_sheet_var, width=30)
        self.p_sheet_combo.grid(row=0, column=1, sticky=tk.W, padx=4)
        self.p_sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_changed)
        ttk.Label(sheets, text="Лист _t:").grid(row=1, column=0, sticky=tk.W)
        self.t_sheet_combo = ttk.Combobox(sheets, textvariable=self.t_sheet_var, width=30)
        self.t_sheet_combo.grid(row=1, column=1, sticky=tk.W, padx=4)
        self.t_sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        columns = ttk.LabelFrame(scroll_frame, text="Колонки", padding=6)
        columns.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        crit_row = ttk.Frame(columns)
        crit_row.pack(fill=tk.X, pady=2)
        ttk.Label(crit_row, text="Критерий (_p):").pack(side=tk.LEFT)
        self.p_criterion_var = tk.StringVar()
        self.p_criterion_combo = ttk.Combobox(
            crit_row, textvariable=self.p_criterion_var, width=30
        )
        self.p_criterion_combo.pack(side=tk.LEFT, padx=6)

        time_row = ttk.Frame(columns)
        time_row.pack(fill=tk.X, pady=2)
        ttk.Label(time_row, text="Time* (_p / _t):").pack(side=tk.LEFT)
        self.p_time_var = tk.StringVar()
        self.t_time_var = tk.StringVar()
        self.p_time_combo = ttk.Combobox(time_row, textvariable=self.p_time_var, width=18)
        self.p_time_combo.pack(side=tk.LEFT, padx=4)
        self.t_time_combo = ttk.Combobox(time_row, textvariable=self.t_time_var, width=18)
        self.t_time_combo.pack(side=tk.LEFT, padx=4)

        prefix_row = ttk.Frame(columns)
        prefix_row.pack(fill=tk.X, pady=2)
        ttk.Label(prefix_row, text="Префикс _p / _t:").pack(side=tk.LEFT)
        self.p_prefix_var = tk.StringVar()
        self.t_prefix_var = tk.StringVar()
        ttk.Entry(prefix_row, textvariable=self.p_prefix_var, width=18).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Entry(prefix_row, textvariable=self.t_prefix_var, width=18).pack(
            side=tk.LEFT, padx=4
        )

        cols_row = ttk.Frame(columns)
        cols_row.pack(fill=tk.BOTH, expand=True, pady=4)
        self.p_col_selector = ColumnSelector(cols_row, "Данные _p (пусто = по префиксу)")
        self.p_col_selector.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.t_col_selector = ColumnSelector(cols_row, "Данные _t (пусто = по префиксу)")
        self.t_col_selector.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.p_col_selector.set_prefix_provider(lambda: self.p_prefix_var.get())
        self.t_col_selector.set_prefix_provider(lambda: self.t_prefix_var.get())

        save_row = ttk.Frame(scroll_frame)
        save_row.pack(pady=8)
        ttk.Button(save_row, text="Сохранить профиль", command=self.save_profile_from_ui).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(save_row, text="Сбросить к умолчанию", command=self.reset_profile_defaults).pack(
            side=tk.LEFT, padx=4
        )

    def _labeled_entry(
        self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(parent, textvariable=variable, width=16).grid(
            row=row, column=1, sticky=tk.W, pady=2
        )

    def _refresh_profile_list(self) -> None:
        names = [p.name for p in get_profiles(self.settings)]
        self.profile_combo["values"] = names
        self.profile_var.set(self.profile.name)

    def _load_profile_to_ui(self, profile: ProcessingProfile) -> None:
        self.profile = profile
        self.profile_var.set(profile.name)
        self.profile_name_var.set(profile.name)
        self.output_format_var.set(profile.output_format)
        self.csv_sep_var.set(profile.csv_separator)
        self.dry_run_report_var.set(profile.write_dry_run_report)
        self.p_suffix_var.set(profile.p_suffix)
        self.t_suffix_var.set(profile.t_suffix)
        self.criterion_mode_var.set(profile.criterion_mode)
        self.time_tolerance_var.set(str(profile.time_tolerance))
        self.max_workers_var.set(str(profile.max_workers))
        self.p_sheet_var.set(str(profile.p_sheet))
        self.t_sheet_var.set(str(profile.t_sheet))
        self.p_criterion_var.set(profile.p_criterion_column)
        self.p_time_var.set(profile.p_time_column)
        self.t_time_var.set(profile.t_time_column)
        self.p_prefix_var.set(profile.p_column_prefix)
        self.t_prefix_var.set(profile.t_column_prefix)
        self.p_col_selector.set_selected(profile.p_data_columns)
        self.t_col_selector.set_selected(profile.t_data_columns)

    def _profile_from_ui(self) -> ProcessingProfile:
        def sheet_value(raw: str) -> int | str:
            raw = raw.strip()
            if raw.isdigit():
                return int(raw)
            return raw

        def parse_float(raw: str, field_name: str) -> float:
            try:
                value = float(raw.strip().replace(",", "."))
            except ValueError as exc:
                raise ValueError(f"Поле «{field_name}» должно быть числом.") from exc
            if value < 0:
                raise ValueError(f"Поле «{field_name}» не может быть отрицательным.")
            return value

        def parse_positive_int(raw: str, field_name: str) -> int:
            try:
                value = int(raw.strip() or "4")
            except ValueError as exc:
                raise ValueError(f"Поле «{field_name}» должно быть целым числом.") from exc
            if value < 1:
                raise ValueError(f"Поле «{field_name}» должно быть не меньше 1.")
            return value

        name = self.profile_name_var.get().strip() or self.profile.name
        return ProcessingProfile(
            name=name,
            p_suffix=self.p_suffix_var.get().strip() or "_p",
            t_suffix=self.t_suffix_var.get().strip() or "_t",
            p_sheet=sheet_value(self.p_sheet_var.get()),
            t_sheet=sheet_value(self.t_sheet_var.get()),
            p_time_column=self.p_time_var.get().strip() or "Time*",
            t_time_column=self.t_time_var.get().strip() or "Time*",
            p_criterion_column=self.p_criterion_var.get().strip() or "pressure_ai0",
            p_data_columns=self.p_col_selector.get_selected(),
            t_data_columns=self.t_col_selector.get_selected(),
            p_column_prefix=self.p_prefix_var.get().strip(),
            t_column_prefix=self.t_prefix_var.get().strip(),
            criterion_mode=self.criterion_mode_var.get() or "closest",  # type: ignore[arg-type]
            time_tolerance=parse_float(self.time_tolerance_var.get(), "Допуск Time*"),
            max_workers=parse_positive_int(self.max_workers_var.get(), "Потоков"),
            output_format=self.output_format_var.get() or "xlsx",  # type: ignore[arg-type]
            csv_separator=self.csv_sep_var.get() or ";",
            write_dry_run_report=self.dry_run_report_var.get(),
        )

    def _profile_from_ui_or_report(self) -> ProcessingProfile | None:
        try:
            return self._profile_from_ui()
        except ValueError as exc:
            self.log_message(f"Ошибка настроек: {exc}")
            messagebox.showerror("Ошибка настроек", str(exc))
            return None

    def on_profile_changed(self, _event=None) -> None:
        name = self.profile_var.get()
        for profile in get_profiles(self.settings):
            if profile.name == name:
                self._load_profile_to_ui(profile)
                break

    def save_profile_from_ui(self) -> None:
        new_profile = self._profile_from_ui_or_report()
        if new_profile is None:
            return
        old_name = self.profile.name
        if new_profile.name != old_name:
            if not rename_profile(self.settings, old_name, new_profile.name):
                if any(p.name == new_profile.name for p in get_profiles(self.settings)):
                    messagebox.showerror(
                        "Ошибка", f"Профиль «{new_profile.name}» уже существует."
                    )
                    return
        self.profile = new_profile
        set_active_profile(self.settings, self.profile)
        save_settings(self.settings)
        self._refresh_profile_list()
        self.log_message(f"Профиль «{self.profile.name}» сохранён.")

    def new_profile(self) -> None:
        base = "Новый профиль"
        names = {p.name for p in get_profiles(self.settings)}
        name = base
        counter = 1
        while name in names:
            counter += 1
            name = f"{base} {counter}"
        profile = ProcessingProfile(name=name)
        set_active_profile(self.settings, profile)
        save_settings(self.settings)
        self._refresh_profile_list()
        self._load_profile_to_ui(profile)
        self.log_message(f"Создан профиль «{name}».")

    def duplicate_profile_ui(self) -> None:
        source = self.profile_var.get()
        names = {p.name for p in get_profiles(self.settings)}
        new_name = f"{source} (копия)"
        counter = 1
        while new_name in names:
            counter += 1
            new_name = f"{source} (копия {counter})"
        copy = duplicate_profile(self.settings, source, new_name)
        if copy:
            save_settings(self.settings)
            self._refresh_profile_list()
            self._load_profile_to_ui(copy)
            self.log_message(f"Скопирован профиль → «{new_name}».")

    def delete_profile_ui(self) -> None:
        name = self.profile_var.get()
        if not messagebox.askyesno("Удалить", f"Удалить профиль «{name}»?"):
            return
        if not delete_profile(self.settings, name):
            messagebox.showwarning("Нельзя удалить", "Нужен хотя бы один профиль.")
            return
        save_settings(self.settings)
        self.profile = get_active_profile(self.settings)
        self._refresh_profile_list()
        self._load_profile_to_ui(self.profile)
        self.log_message(f"Профиль «{name}» удалён.")

    def import_profile_ui(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("JSON профиль", "*.json"), ("Все файлы", "*.*")]
        )
        if not path:
            return
        try:
            profile = import_profile_from_file(Path(path))
            names = {p.name for p in get_profiles(self.settings)}
            if profile.name in names:
                profile.name = f"{profile.name} (импорт)"
            set_active_profile(self.settings, profile)
            save_settings(self.settings)
            self._refresh_profile_list()
            self._load_profile_to_ui(profile)
            self.log_message(f"Импортирован профиль «{profile.name}».")
        except Exception as exc:
            messagebox.showerror("Ошибка импорта", str(exc))

    def export_profile_ui(self) -> None:
        profile = self._profile_from_ui_or_report()
        if profile is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=f"{profile.name}.json",
            filetypes=[("JSON профиль", "*.json")],
        )
        if not path:
            return
        try:
            export_profile_to_file(profile, Path(path))
            self.log_message(f"Профиль экспортирован: {path}")
        except Exception as exc:
            messagebox.showerror("Ошибка экспорта", str(exc))

    def reset_profile_defaults(self) -> None:
        if not messagebox.askyesno("Сброс", "Вернуть поля к значениям по умолчанию?"):
            return
        defaults = ProcessingProfile(name=self.profile_name_var.get())
        self._load_profile_to_ui(defaults)

    def validate_before_run(self) -> None:
        folder = self.folder_path.get().strip()
        if not folder:
            self.log_message("Укажите папку для проверки.")
            return
        profile = self._profile_from_ui_or_report()
        if profile is None:
            return
        target = None
        target_str = self.target_value.get().strip()
        if target_str:
            try:
                target = float(target_str.replace(",", "."))
            except ValueError:
                self.log_message("Целевое значение не число — проверка критерия пропущена.")

        errors, warnings = validate_setup(Path(folder), profile, target)
        self.log_area.delete(1.0, tk.END)
        if errors:
            self.log_message("ПРОВЕРКА: найдены ошибки:")
            for err in errors:
                self.log_message(f"  ✗ {err}")
        else:
            self.log_message("ПРОВЕРКА: критических ошибок нет.")
        if warnings:
            self.log_message("Предупреждения:")
            for warn in warnings:
                self.log_message(f"  ! {warn}")
        pairs = discover_pairs(Path(folder), profile.p_suffix, profile.t_suffix)
        self.log_message(f"Полных пар: {len(pairs)}")

    def browse_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def scan_folder(self) -> None:
        folder = self.folder_path.get().strip()
        if not folder:
            self.log_message("Укажите папку для сканирования.")
            return
        folder_path = Path(folder)
        pairs = discover_pairs(
            folder_path,
            self.p_suffix_var.get().strip() or "_p",
            self.t_suffix_var.get().strip() or "_t",
        )
        if not pairs:
            self.log_message("Пары _p/_t не найдены.")
            return

        sample_p = pairs[0].p_path
        sample_t = pairs[0].t_path
        self._scan_data["p"] = scan_file(sample_p)
        self._scan_data["t"] = scan_file(sample_t)

        p_sheets = list(self._scan_data["p"].keys())
        t_sheets = list(self._scan_data["t"].keys())
        self.p_sheet_combo["values"] = p_sheets
        self.t_sheet_combo["values"] = t_sheets
        self._select_sheet_combo(self.p_sheet_combo, self.p_sheet_var, p_sheets, self.profile.p_sheet)
        self._select_sheet_combo(self.t_sheet_combo, self.t_sheet_var, t_sheets, self.profile.t_sheet)

        self._update_columns_from_scan()
        self.log_message(
            f"Сканирование: {len(pairs)} пар. Образцы: {sample_p.name}, {sample_t.name}"
        )

    def _update_columns_from_scan(self) -> None:
        p_sheet = self.p_sheet_var.get()
        t_sheet = self.t_sheet_var.get()
        p_cols = self._scan_data.get("p", {}).get(p_sheet, [])
        t_cols = self._scan_data.get("t", {}).get(t_sheet, [])
        self.p_col_selector.set_available(p_cols)
        self.t_col_selector.set_available(t_cols)
        self.p_criterion_combo["values"] = p_cols
        self.p_time_combo["values"] = p_cols
        self.t_time_combo["values"] = t_cols

    def _select_sheet_combo(
        self,
        combo: ttk.Combobox,
        variable: tk.StringVar,
        sheets: list[str],
        preferred: int | str,
    ) -> None:
        if not sheets:
            return
        if isinstance(preferred, int) and 0 <= preferred < len(sheets):
            variable.set(sheets[preferred])
        elif str(preferred) in sheets:
            variable.set(str(preferred))
        else:
            variable.set(sheets[min(1, len(sheets) - 1)])

    def _on_sheet_changed(self, _event=None) -> None:
        self._update_columns_from_scan()

    def log_message(self, msg: str) -> None:
        def append() -> None:
            self.log_area.insert(tk.END, msg + "\n")
            self.log_area.see(tk.END)
            logging.info(msg)

        self.root.after(0, append)

    def update_progress(self, current: int, total: int, label: str) -> None:
        def update() -> None:
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = current
            self.status_var.set(f"Обработка: {current}/{total} — {label}")

        self.root.after(0, update)

    def start_processing(self) -> None:
        if self.processing_thread and self.processing_thread.is_alive():
            self.log_message("Обработка уже выполняется.")
            return

        folder = self.folder_path.get().strip()
        if not folder:
            self.log_message("Ошибка: выберите папку.")
            return

        target_str = self.target_value.get().strip()
        if not target_str:
            self.log_message("Ошибка: введите целевое значение.")
            return
        try:
            target = float(target_str.replace(",", "."))
        except ValueError:
            self.log_message("Ошибка: целевое значение должно быть числом.")
            return

        profile = self._profile_from_ui_or_report()
        if profile is None:
            return
        self.profile = profile
        set_active_profile(self.settings, self.profile)
        self.settings["last_folder"] = folder
        self.settings["last_target"] = target_str
        save_settings(self.settings)

        self.cancel_event.clear()
        self.log_area.delete(1.0, tk.END)
        self.log_message("Начинаю обработку...")
        self.progress["value"] = 0
        self.run_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)
        self.status_var.set("Обработка...")

        self.processing_thread = threading.Thread(
            target=run_processing,
            args=(
                folder,
                target,
                self.profile,
                self.log_message,
                self.processing_done,
                self.update_progress,
                self.cancel_event,
            ),
            daemon=True,
        )
        self.processing_thread.start()

    def cancel_processing(self) -> None:
        self.cancel_event.set()
        self.log_message("Запрошена отмена...")

    def processing_done(self, report: ProcessingReport) -> None:
        def finish() -> None:
            self.run_btn.configure(state=tk.NORMAL)
            self.cancel_btn.configure(state=tk.DISABLED)
            self.progress["value"] = 0
            if report.success:
                if report.output_paths:
                    self.status_var.set(f"Готово: {report.output_paths[0]}")
                    files = "\n  ".join(report.output_paths)
                    self.log_message(
                        f"\nУСПЕШНО. Обработано: {report.processed}, "
                        f"пропущено: {report.skipped}.\n  {files}"
                    )
                else:
                    self.status_var.set("Готово (DRY-RUN)")
                    msg = (
                        f"\nУСПЕШНО (DRY-RUN). Обработано: {report.processed}, "
                        f"пропущено: {report.skipped}"
                    )
                    if report.dry_run_report_path:
                        msg += f"\n  Отчёт: {report.dry_run_report_path}"
                    self.log_message(msg)
            else:
                self.status_var.set("Ошибка (см. лог)")
                self.log_message(
                    f"\nОБРАБОТКА ЗАВЕРШЕНА С ОШИБКОЙ. "
                    f"Обработано: {report.processed}, пропущено: {report.skipped}"
                )
            self.processing_thread = None

        self.root.after(0, finish)

    def on_closing(self) -> None:
        if self.processing_thread and self.processing_thread.is_alive():
            messagebox.showwarning(
                "Подождите", "Дождитесь завершения обработки или нажмите «Отмена»."
            )
            return
        self.settings["last_folder"] = self.folder_path.get()
        self.settings["last_target"] = self.target_value.get()
        save_settings(self.settings)
        self.root.destroy()
