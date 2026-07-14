from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from arenda_table import TaxRateRule, build_arenda_workbook, save_arenda_workbook
from depreciation_table import (
    MONTH_NAMES,
    MONTH_TO_INDEX,
    ContractInput,
    Period,
    generate_depreciation_workbook,
    list_table_names,
    save_depreciation_workbook,
)


@dataclass
class LessorInput:
    index: int
    frame: ttk.LabelFrame
    lessor_name_var: tk.StringVar
    files: list[Path] = field(default_factory=list)
    files_var: tk.StringVar = field(default_factory=tk.StringVar)


@dataclass
class ContractFileInput:
    frame: ttk.LabelFrame
    file_label_var: tk.StringVar
    number_var: tk.StringVar
    file_path: Path | None = None


@dataclass
class TaxRuleInput:
    frame: ttk.Frame
    year_var: tk.StringVar
    multiplier_var: tk.StringVar


class ArendaUI:
    SOURCE_EXISTING = "existing"
    SOURCE_NEW_IN_FILE = "new_in_file"
    SOURCE_NEW_FILE = "new_file"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Генератор таблиц")
        self.root.geometry("1040x760")

        self.task_var = tk.StringVar(value="")
        self.tenant_var = tk.StringVar()

        self.lessor_items: list[LessorInput] = []
        self.tax_rule_items: list[TaxRuleInput] = []

        self.source_mode_var = tk.StringVar(value=self.SOURCE_EXISTING)
        self.table_file: Path | None = None
        self.table_file_var = tk.StringVar(value="Файл не выбран")
        self.table_name_var = tk.StringVar()
        self.new_table_name_var = tk.StringVar()
        self.detected_table_names: list[str] = []

        self.statement_file: Path | None = None
        self.statement_file_var = tk.StringVar(value="Файл не выбран")

        today = date.today()
        self.month_var = tk.StringVar(value=MONTH_NAMES[today.month - 1].capitalize())
        self.year_var = tk.StringVar(value=str(today.year))

        self.contract_items: list[ContractFileInput] = []

        self.main = ttk.Frame(self.root, padding=16)
        self.main.pack(fill="both", expand=True)

        self._build_header()
        self._build_task_container()

    def _build_header(self) -> None:
        header = ttk.Label(self.main, text="Инструменты бухгалтера", font=("Helvetica", 16, "bold"))
        header.grid(row=0, column=0, sticky="w")

        task_row = ttk.Frame(self.main)
        task_row.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        task_row.columnconfigure(1, weight=1)

        ttk.Label(task_row, text="Задача").grid(row=0, column=0, sticky="w", padx=(0, 8))

        task_combo = ttk.Combobox(
            task_row,
            textvariable=self.task_var,
            state="readonly",
            values=["Аренда", "Амортизация"],
        )
        task_combo.grid(row=0, column=1, sticky="ew")
        task_combo.bind("<<ComboboxSelected>>", self._on_task_changed)

    def _build_task_container(self) -> None:
        self.task_canvas = tk.Canvas(self.main, borderwidth=0, highlightthickness=0)
        self.task_canvas.grid(row=2, column=0, sticky="nsew")

        self.task_scrollbar = ttk.Scrollbar(self.main, orient="vertical", command=self.task_canvas.yview)
        self.task_scrollbar.grid(row=2, column=1, sticky="ns")
        self.task_canvas.configure(yscrollcommand=self.task_scrollbar.set)

        self.task_container = ttk.Frame(self.task_canvas)
        self.task_container.columnconfigure(0, weight=1)
        self.task_canvas_window = self.task_canvas.create_window((0, 0), window=self.task_container, anchor="nw")

        self.task_container.bind("<Configure>", self._refresh_task_scrollregion)
        self.task_canvas.bind("<Configure>", self._resize_task_body_to_canvas)

        self.root.bind_all("<MouseWheel>", self._on_task_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_task_mousewheel_linux, add="+")
        self.root.bind_all("<Button-5>", self._on_task_mousewheel_linux, add="+")

        self.main.rowconfigure(2, weight=1)
        self.main.columnconfigure(0, weight=1)

        self.arenda_panel = ttk.Frame(self.task_container)
        self._build_arenda_panel(self.arenda_panel)

        self.amortization_panel = ttk.Frame(self.task_container)
        self._build_amortization_panel(self.amortization_panel)

    def _refresh_task_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.task_canvas.configure(scrollregion=self.task_canvas.bbox("all"))

    def _resize_task_body_to_canvas(self, event: tk.Event) -> None:
        self.task_canvas.itemconfigure(self.task_canvas_window, width=event.width)

    def _widget_inside_task_area(self, widget: tk.Misc | None) -> bool:
        while widget is not None:
            if widget is self.task_canvas or widget is self.task_container:
                return True
            widget = widget.master
        return False

    def _pointer_inside_task_area(self) -> bool:
        widget = self.root.winfo_containing(self.root.winfo_pointerx(), self.root.winfo_pointery())
        return self._widget_inside_task_area(widget)

    def _on_task_mousewheel(self, event: tk.Event) -> None:
        if not self._pointer_inside_task_area() or self._pointer_inside_lessor_list() or self._pointer_inside_contracts_list():
            return

        delta = getattr(event, "delta", 0)
        if not delta:
            return

        if abs(delta) < 120:
            units = -1 if delta > 0 else 1
        else:
            units = int(-delta / 120)

        if units:
            self.task_canvas.yview_scroll(units, "units")

    def _on_task_mousewheel_linux(self, event: tk.Event) -> None:
        if not self._pointer_inside_task_area() or self._pointer_inside_lessor_list() or self._pointer_inside_contracts_list():
            return

        button = getattr(event, "num", None)
        if button == 4:
            self.task_canvas.yview_scroll(-1, "units")
        elif button == 5:
            self.task_canvas.yview_scroll(1, "units")

    def _build_arenda_panel(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        tenant_box = ttk.LabelFrame(parent, text="Параметры арендатора", padding=10)
        tenant_box.grid(row=0, column=0, sticky="ew")
        tenant_box.columnconfigure(1, weight=1)

        ttk.Label(tenant_box, text="Арендатор").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(tenant_box, textvariable=self.tenant_var).grid(row=0, column=1, sticky="ew")
        ttk.Label(
            tenant_box,
            text="Название компании должно точно совпадать с названием в таблицах анализа.",
            foreground="#4a4a4a",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        lessors_box = ttk.LabelFrame(parent, text="Арендодатели и файлы", padding=10)
        lessors_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        lessors_box.columnconfigure(0, weight=1)

        lessor_actions = ttk.Frame(lessors_box)
        lessor_actions.grid(row=0, column=0, sticky="ew")
        ttk.Button(lessor_actions, text="Добавить арендодателя", command=self._add_lessor).pack(side="left")

        list_frame = ttk.Frame(lessors_box)
        list_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        list_frame.columnconfigure(0, weight=1)

        self.lessors_canvas = tk.Canvas(list_frame, height=260, borderwidth=0, highlightthickness=0)
        self.lessors_canvas.grid(row=0, column=0, sticky="ew")

        self.lessors_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.lessors_canvas.yview)
        self.lessors_scrollbar.grid(row=0, column=1, sticky="ns")
        self.lessors_canvas.configure(yscrollcommand=self.lessors_scrollbar.set)

        self.lessors_body = ttk.Frame(self.lessors_canvas)
        self.lessors_body.columnconfigure(0, weight=1)
        self.lessors_canvas_window = self.lessors_canvas.create_window((0, 0), window=self.lessors_body, anchor="nw")
        self.lessors_body.bind("<Configure>", self._refresh_lessor_scrollregion)
        self.lessors_canvas.bind("<Configure>", self._resize_lessor_body_to_canvas)
        self.root.bind_all("<MouseWheel>", self._on_lessor_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_lessor_mousewheel_linux, add="+")
        self.root.bind_all("<Button-5>", self._on_lessor_mousewheel_linux, add="+")

        taxes_box = ttk.LabelFrame(parent, text="Налоговые множители", padding=10)
        taxes_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        taxes_box.columnconfigure(0, weight=1)

        ttk.Label(taxes_box, text="Правила применяются по году: с указанного года и далее.").grid(row=0, column=0, sticky="w")

        self.tax_rules_frame = ttk.Frame(taxes_box)
        self.tax_rules_frame.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.tax_rules_frame.columnconfigure(0, weight=1)

        ttk.Button(taxes_box, text="Добавить правило", command=self._add_tax_rule).grid(row=2, column=0, sticky="w", pady=(8, 0))

        actions = ttk.Frame(parent)
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        actions.columnconfigure(0, weight=1)

        ttk.Button(actions, text="Сформировать файл Аренда", command=self._generate_arenda).grid(row=0, column=1, sticky="e")

        self._add_tax_rule("0", "0.8")
        self._add_tax_rule("2025", "0.75")
        self._add_lessor()

    def _on_task_changed(self, _event: tk.Event | None = None) -> None:
        for child in self.task_container.winfo_children():
            child.grid_forget()

        if self.task_var.get() == "Аренда":
            self.arenda_panel.grid(row=0, column=0, sticky="nsew")
            self.task_container.rowconfigure(0, weight=1)
            self.task_container.columnconfigure(0, weight=1)
        elif self.task_var.get() == "Амортизация":
            self.amortization_panel.grid(row=0, column=0, sticky="nsew")
            self.task_container.rowconfigure(0, weight=1)
            self.task_container.columnconfigure(0, weight=1)

    def _refresh_lessor_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.lessors_canvas.configure(scrollregion=self.lessors_canvas.bbox("all"))

    def _resize_lessor_body_to_canvas(self, event: tk.Event) -> None:
        self.lessors_canvas.itemconfigure(self.lessors_canvas_window, width=event.width)

    def _widget_inside_lessor_list(self, widget: tk.Misc | None) -> bool:
        while widget is not None:
            if widget is self.lessors_canvas or widget is self.lessors_body:
                return True
            widget = widget.master
        return False

    def _pointer_inside_lessor_list(self) -> bool:
        widget = self.root.winfo_containing(self.root.winfo_pointerx(), self.root.winfo_pointery())
        return self._widget_inside_lessor_list(widget)

    def _on_lessor_mousewheel(self, event: tk.Event) -> None:
        if not self._pointer_inside_lessor_list():
            return

        delta = getattr(event, "delta", 0)
        if not delta:
            return

        # On macOS, delta is often small (e.g., +/-1..10), not multiples of 120.
        if abs(delta) < 120:
            units = -1 if delta > 0 else 1
        else:
            units = int(-delta / 120)

        if units:
            self.lessors_canvas.yview_scroll(units, "units")

    def _on_lessor_mousewheel_linux(self, event: tk.Event) -> None:
        if not self._pointer_inside_lessor_list():
            return

        button = getattr(event, "num", None)
        if button == 4:
            self.lessors_canvas.yview_scroll(-1, "units")
        elif button == 5:
            self.lessors_canvas.yview_scroll(1, "units")

    def _add_lessor(self) -> None:
        idx = len(self.lessor_items) + 1
        wrapper = ttk.LabelFrame(self.lessors_body, text=f"Арендодатель {idx}", padding=8)
        wrapper.grid(row=len(self.lessor_items), column=0, sticky="ew", pady=(0, 8))
        wrapper.columnconfigure(1, weight=1)

        name_var = tk.StringVar()
        ttk.Label(wrapper, text="Название").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(wrapper, textvariable=name_var).grid(row=0, column=1, sticky="ew")

        files_var = tk.StringVar(value="Файлы не выбраны")
        files_label = ttk.Label(wrapper, textvariable=files_var, justify="left")
        files_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 6))

        item = LessorInput(index=idx, frame=wrapper, lessor_name_var=name_var, files_var=files_var)
        self.lessor_items.append(item)

        actions = ttk.Frame(wrapper)
        actions.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Button(actions, text="Выбрать файлы", command=lambda target=item: self._choose_files_for_lessor(target)).pack(side="left")
        ttk.Button(actions, text="Очистить", command=lambda target=item: self._clear_lessor_files(target)).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Удалить", command=lambda target=item: self._remove_lessor(target)).pack(side="left", padx=(8, 0))

        self._refresh_lessor_scrollregion()

    def _remove_lessor(self, item: LessorInput) -> None:
        if len(self.lessor_items) == 1:
            messagebox.showwarning("Нельзя удалить", "Должен остаться хотя бы один арендодатель.")
            return

        item.frame.destroy()
        self.lessor_items = [entry for entry in self.lessor_items if entry is not item]

        for idx, entry in enumerate(self.lessor_items, start=1):
            entry.index = idx
            entry.frame.configure(text=f"Арендодатель {idx}")
            entry.frame.grid_configure(row=idx - 1)

        self._refresh_lessor_scrollregion()

    def _choose_files_for_lessor(self, item: LessorInput) -> None:
        chosen = filedialog.askopenfilenames(
            title="Выберите файлы анализа",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            self._set_lessor_files(item, [Path(path) for path in chosen])

    def _clear_lessor_files(self, item: LessorInput) -> None:
        item.files = []
        item.files_var.set("Файлы не выбраны")

    def _set_lessor_files(self, item: LessorInput, files: list[Path]) -> None:
        unique: list[Path] = []
        seen: set[Path] = set()
        for file_path in files:
            resolved = Path(file_path)
            if resolved not in seen:
                seen.add(resolved)
                unique.append(resolved)
        item.files = unique

        if not item.files:
            item.files_var.set("Файлы не выбраны")
            return

        preview = "\n".join(str(path) for path in item.files)
        item.files_var.set(preview)

    def _add_tax_rule(self, year: str = "", multiplier: str = "") -> None:
        row = len(self.tax_rule_items)
        frame = ttk.Frame(self.tax_rules_frame)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        year_var = tk.StringVar(value=year)
        multiplier_var = tk.StringVar(value=multiplier)

        ttk.Label(frame, text="С года").grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="Множитель").grid(row=0, column=1, sticky="w")

        ttk.Entry(frame, textvariable=year_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Entry(frame, textvariable=multiplier_var).grid(row=1, column=1, sticky="ew", padx=(0, 8))

        item = TaxRuleInput(frame=frame, year_var=year_var, multiplier_var=multiplier_var)
        self.tax_rule_items.append(item)

        ttk.Button(frame, text="Удалить", command=lambda target=item: self._remove_tax_rule(target)).grid(row=1, column=2, sticky="e")

    def _remove_tax_rule(self, item: TaxRuleInput) -> None:
        if len(self.tax_rule_items) == 1:
            messagebox.showwarning("Нельзя удалить", "Должно остаться хотя бы одно налоговое правило.")
            return

        item.frame.destroy()
        self.tax_rule_items = [entry for entry in self.tax_rule_items if entry is not item]

        for row, entry in enumerate(self.tax_rule_items):
            entry.frame.grid_configure(row=row)

    def _parse_tax_rules(self) -> list[TaxRateRule]:
        parsed: list[TaxRateRule] = []
        for item in self.tax_rule_items:
            year_text = item.year_var.get().strip()
            multiplier_text = item.multiplier_var.get().strip().replace(",", ".")

            if not year_text and not multiplier_text:
                continue
            if not year_text or not multiplier_text:
                raise ValueError("Укажите и год, и множитель в каждом налоговом правиле.")

            year = int(year_text)
            multiplier = float(multiplier_text)
            parsed.append(TaxRateRule(from_year=year, multiplier=multiplier))

        if not parsed:
            raise ValueError("Добавьте хотя бы одно налоговое правило.")

        parsed.sort(key=lambda rule: rule.from_year)
        return parsed

    def _collect_lessor_groups(self) -> dict[str, list[Path]]:
        groups: dict[str, list[Path]] = {}
        for item in self.lessor_items:
            lessor_name = item.lessor_name_var.get().strip()
            if not lessor_name and not item.files:
                continue
            if not lessor_name:
                raise ValueError("У каждого арендодателя с файлами должно быть название.")
            if not item.files:
                raise ValueError(f"Для арендодателя '{lessor_name}' нужно выбрать хотя бы один файл.")
            groups[lessor_name] = item.files

        if not groups:
            raise ValueError("Добавьте хотя бы одного арендодателя с файлами.")

        return groups

    def _generate_arenda(self) -> None:
        try:
            tenant = self.tenant_var.get().strip()
            if not tenant:
                raise ValueError("Заполните поле 'Арендатор'.")

            lessor_groups = self._collect_lessor_groups()
            tax_rules = self._parse_tax_rules()

            output_path = filedialog.asksaveasfilename(
                title="Куда сохранить файл",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile="Амортизация_аренда.xlsx",
            )
            if not output_path:
                return

            workbook = build_arenda_workbook(
                lessor_groups=lessor_groups,
                lessee_company_name=tenant,
                tax_rules=tax_rules,
            )
            save_arenda_workbook(workbook, output_path)
            messagebox.showinfo("Готово", f"Файл сохранен:\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    def _build_amortization_panel(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)

        source_box = ttk.LabelFrame(parent, text="Итоговая таблица Амортизации", padding=10)
        source_box.grid(row=0, column=0, sticky="ew")
        source_box.columnconfigure(0, weight=1)

        ttk.Radiobutton(
            source_box,
            text="Дополнить существующую таблицу новым месяцем/договором",
            variable=self.source_mode_var,
            value=self.SOURCE_EXISTING,
            command=self._on_source_mode_changed,
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            source_box,
            text="Создать новую таблицу в загруженном файле",
            variable=self.source_mode_var,
            value=self.SOURCE_NEW_IN_FILE,
            command=self._on_source_mode_changed,
        ).grid(row=1, column=0, sticky="w")
        ttk.Radiobutton(
            source_box,
            text="Создать новую таблицу в новом файле",
            variable=self.source_mode_var,
            value=self.SOURCE_NEW_FILE,
            command=self._on_source_mode_changed,
        ).grid(row=2, column=0, sticky="w")

        self.table_file_row = ttk.Frame(source_box)
        self.table_file_row.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.table_file_row.columnconfigure(1, weight=1)
        ttk.Button(self.table_file_row, text="Выбрать файл таблицы", command=self._choose_table_file).grid(row=0, column=0, sticky="w")
        ttk.Label(self.table_file_row, textvariable=self.table_file_var, justify="left").grid(row=0, column=1, sticky="ew", padx=(8, 0))

        self.table_name_row = ttk.Frame(source_box)
        self.table_name_row.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(self.table_name_row, text="Название таблицы").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.table_name_combo = ttk.Combobox(self.table_name_row, textvariable=self.table_name_var, width=30)
        self.table_name_combo.grid(row=0, column=1, sticky="w")

        self.new_table_row = ttk.Frame(source_box)
        self.new_table_row.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(self.new_table_row, text="Название новой таблицы").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(self.new_table_row, textvariable=self.new_table_name_var, width=32).grid(row=0, column=1, sticky="w")

        statement_box = ttk.LabelFrame(parent, text="Ведомость Амортизации", padding=10)
        statement_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        statement_box.columnconfigure(1, weight=1)
        ttk.Button(statement_box, text="Выбрать файл", command=self._choose_statement_file).grid(row=0, column=0, sticky="w")
        ttk.Label(statement_box, textvariable=self.statement_file_var, justify="left").grid(row=0, column=1, sticky="ew", padx=(8, 0))

        contracts_box = ttk.LabelFrame(parent, text="Договоры Аренды", padding=10)
        contracts_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        contracts_box.columnconfigure(0, weight=1)

        contracts_actions = ttk.Frame(contracts_box)
        contracts_actions.grid(row=0, column=0, sticky="ew")
        ttk.Button(contracts_actions, text="Добавить договор", command=self._add_contract).pack(side="left")

        list_frame = ttk.Frame(contracts_box)
        list_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        list_frame.columnconfigure(0, weight=1)

        self.contracts_canvas = tk.Canvas(list_frame, height=220, borderwidth=0, highlightthickness=0)
        self.contracts_canvas.grid(row=0, column=0, sticky="ew")

        self.contracts_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.contracts_canvas.yview)
        self.contracts_scrollbar.grid(row=0, column=1, sticky="ns")
        self.contracts_canvas.configure(yscrollcommand=self.contracts_scrollbar.set)

        self.contracts_body = ttk.Frame(self.contracts_canvas)
        self.contracts_body.columnconfigure(0, weight=1)
        self.contracts_canvas_window = self.contracts_canvas.create_window((0, 0), window=self.contracts_body, anchor="nw")
        self.contracts_body.bind("<Configure>", self._refresh_contracts_scrollregion)
        self.contracts_canvas.bind("<Configure>", self._resize_contracts_body_to_canvas)
        self.root.bind_all("<MouseWheel>", self._on_contracts_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_contracts_mousewheel_linux, add="+")
        self.root.bind_all("<Button-5>", self._on_contracts_mousewheel_linux, add="+")

        ttk.Label(
            parent,
            text=(
                "Ведомость Амортизации и все файлы Договоров Аренды должны относиться "
                "к одному и тому же месяцу и году - тому, что выбран ниже."
            ),
            foreground="#4a4a4a",
            wraplength=820,
            justify="left",
        ).grid(row=3, column=0, sticky="ew", pady=(12, 0))

        period_box = ttk.LabelFrame(parent, text="Период", padding=10)
        period_box.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(period_box, text="Месяц").grid(row=0, column=0, sticky="w", padx=(0, 8))
        month_combo = ttk.Combobox(
            period_box,
            textvariable=self.month_var,
            state="readonly",
            values=[name.capitalize() for name in MONTH_NAMES],
            width=14,
        )
        month_combo.grid(row=0, column=1, sticky="w")
        ttk.Label(period_box, text="Год").grid(row=0, column=2, sticky="w", padx=(16, 8))
        ttk.Entry(period_box, textvariable=self.year_var, width=8).grid(row=0, column=3, sticky="w")

        actions = ttk.Frame(parent)
        actions.grid(row=5, column=0, sticky="ew", pady=(16, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Button(actions, text="Сформировать", command=self._generate_amortization).grid(row=0, column=1, sticky="e")

        self._on_source_mode_changed()
        self._add_contract()

    def _on_source_mode_changed(self, _event: tk.Event | None = None) -> None:
        mode = self.source_mode_var.get()
        show_file = mode in (self.SOURCE_EXISTING, self.SOURCE_NEW_IN_FILE)
        show_new_name = mode in (self.SOURCE_NEW_IN_FILE, self.SOURCE_NEW_FILE)

        if show_file:
            self.table_file_row.grid()
        else:
            self.table_file_row.grid_remove()
            self.table_file = None
            self.table_file_var.set("Файл не выбран")
            self.detected_table_names = []

        if show_new_name:
            self.new_table_row.grid()
        else:
            self.new_table_row.grid_remove()

        self._refresh_table_name_visibility()

    def _refresh_table_name_visibility(self) -> None:
        show = self.source_mode_var.get() == self.SOURCE_EXISTING and len(self.detected_table_names) > 1
        if show:
            self.table_name_combo["values"] = self.detected_table_names
            self.table_name_row.grid()
        else:
            self.table_name_row.grid_remove()

    def _choose_table_file(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Выберите файл итоговой таблицы Амортизации",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if not chosen:
            return

        path = Path(chosen)
        if path.suffix.lower() != ".xlsx":
            messagebox.showerror("Ошибка", "Файл итоговой таблицы должен быть в формате .xlsx.")
            return

        self.table_file = path
        self.table_file_var.set(str(path))
        try:
            self.detected_table_names = list_table_names(path)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{exc}")
            self.detected_table_names = []

        if len(self.detected_table_names) == 1:
            self.table_name_var.set(self.detected_table_names[0])
        else:
            self.table_name_var.set("")

        self._refresh_table_name_visibility()

    def _choose_statement_file(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Выберите файл Ведомости Амортизации",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            self.statement_file = Path(chosen)
            self.statement_file_var.set(str(self.statement_file))

    def _refresh_contracts_scrollregion(self, _event: tk.Event | None = None) -> None:
        self.contracts_canvas.configure(scrollregion=self.contracts_canvas.bbox("all"))

    def _resize_contracts_body_to_canvas(self, event: tk.Event) -> None:
        self.contracts_canvas.itemconfigure(self.contracts_canvas_window, width=event.width)

    def _widget_inside_contracts_list(self, widget: tk.Misc | None) -> bool:
        while widget is not None:
            if widget is self.contracts_canvas or widget is self.contracts_body:
                return True
            widget = widget.master
        return False

    def _pointer_inside_contracts_list(self) -> bool:
        widget = self.root.winfo_containing(self.root.winfo_pointerx(), self.root.winfo_pointery())
        return self._widget_inside_contracts_list(widget)

    def _on_contracts_mousewheel(self, event: tk.Event) -> None:
        if not self._pointer_inside_contracts_list():
            return

        delta = getattr(event, "delta", 0)
        if not delta:
            return

        if abs(delta) < 120:
            units = -1 if delta > 0 else 1
        else:
            units = int(-delta / 120)

        if units:
            self.contracts_canvas.yview_scroll(units, "units")

    def _on_contracts_mousewheel_linux(self, event: tk.Event) -> None:
        if not self._pointer_inside_contracts_list():
            return

        button = getattr(event, "num", None)
        if button == 4:
            self.contracts_canvas.yview_scroll(-1, "units")
        elif button == 5:
            self.contracts_canvas.yview_scroll(1, "units")

    def _add_contract(self) -> None:
        idx = len(self.contract_items) + 1
        wrapper = ttk.LabelFrame(self.contracts_body, text=f"Договор {idx}", padding=8)
        wrapper.grid(row=len(self.contract_items), column=0, sticky="ew", pady=(0, 8))
        wrapper.columnconfigure(1, weight=1)

        file_var = tk.StringVar(value="Файл не выбран")
        ttk.Label(wrapper, textvariable=file_var, justify="left").grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(wrapper, text="Номер договора").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
        number_var = tk.StringVar()
        ttk.Entry(wrapper, textvariable=number_var).grid(row=1, column=1, sticky="ew", pady=(6, 0))

        item = ContractFileInput(frame=wrapper, file_label_var=file_var, number_var=number_var)
        self.contract_items.append(item)

        actions = ttk.Frame(wrapper)
        actions.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(actions, text="Выбрать файл", command=lambda target=item: self._choose_contract_file(target)).pack(side="left")
        ttk.Button(actions, text="Удалить", command=lambda target=item: self._remove_contract(target)).pack(side="left", padx=(8, 0))

        self._refresh_contracts_scrollregion()

    def _choose_contract_file(self, item: ContractFileInput) -> None:
        chosen = filedialog.askopenfilename(
            title="Выберите файл договора аренды",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            item.file_path = Path(chosen)
            item.file_label_var.set(str(item.file_path))

    def _remove_contract(self, item: ContractFileInput) -> None:
        if len(self.contract_items) == 1:
            messagebox.showwarning("Нельзя удалить", "Должен остаться хотя бы один договор.")
            return

        item.frame.destroy()
        self.contract_items = [entry for entry in self.contract_items if entry is not item]

        for idx, entry in enumerate(self.contract_items, start=1):
            entry.frame.configure(text=f"Договор {idx}")
            entry.frame.grid_configure(row=idx - 1)

        self._refresh_contracts_scrollregion()

    def _collect_contracts(self) -> list[ContractInput]:
        contracts: list[ContractInput] = []
        seen_numbers: set[str] = set()
        for item in self.contract_items:
            number = item.number_var.get().strip()
            if item.file_path is None and not number:
                continue
            if item.file_path is None:
                raise ValueError("У каждого договора должен быть выбран файл.")
            if not number:
                raise ValueError("У каждого договора должен быть указан номер.")
            if number in seen_numbers:
                raise ValueError(f"Номер договора '{number}' указан более одного раза.")
            seen_numbers.add(number)
            contracts.append(ContractInput(contract_number=number, contract_file=item.file_path))

        if not contracts:
            raise ValueError("Добавьте хотя бы один договор аренды.")

        return contracts

    def _generate_amortization(self) -> None:
        try:
            month_text = self.month_var.get().strip().lower()
            if month_text not in MONTH_TO_INDEX:
                raise ValueError("Выберите месяц периода.")

            year_text = self.year_var.get().strip()
            if not year_text.isdigit():
                raise ValueError("Укажите год периода числом.")

            period = Period(year=int(year_text), month_index=MONTH_TO_INDEX[month_text])

            if not self.statement_file:
                raise ValueError("Выберите файл 'Ведомость Амортизации'.")

            contracts = self._collect_contracts()

            mode = self.source_mode_var.get()
            source_table_path: Path | None = None
            source_table_name: str | None = None
            new_table_name: str | None = None

            if mode == self.SOURCE_EXISTING:
                if not self.table_file:
                    raise ValueError("Выберите файл итоговой таблицы Амортизации.")
                source_table_path = self.table_file
                source_table_name = self.table_name_var.get().strip()
                if not source_table_name:
                    raise ValueError("Укажите название таблицы, которую нужно дополнить.")
            elif mode == self.SOURCE_NEW_IN_FILE:
                if not self.table_file:
                    raise ValueError("Выберите файл, в который нужно добавить новую таблицу.")
                source_table_path = self.table_file
                new_table_name = self.new_table_name_var.get().strip()
                if not new_table_name:
                    raise ValueError("Укажите название новой таблицы.")
            else:
                new_table_name = self.new_table_name_var.get().strip()
                if not new_table_name:
                    raise ValueError("Укажите название новой таблицы.")

            result = generate_depreciation_workbook(
                period=period,
                statement_file=self.statement_file,
                contracts=contracts,
                source_table_path=source_table_path,
                source_table_name=source_table_name,
                new_table_name=new_table_name,
            )

            output_path = filedialog.asksaveasfilename(
                title="Куда сохранить файл",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile="Амортизация.xlsx",
            )
            if not output_path:
                return

            save_depreciation_workbook(result, output_path)
            messagebox.showinfo("Готово", "Файл сохранен:\n{}\n\n{}".format(output_path, "\n".join(result.messages)))
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))


def create_root() -> tk.Tk:
    return tk.Tk()


def main() -> None:
    root = create_root()
    app = ArendaUI(root)
    app.task_var.set("Аренда")
    app._on_task_changed()
    root.mainloop()


if __name__ == "__main__":
    main()
