from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import openpyxl
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


MONTH_NAMES = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]
MONTH_TO_INDEX = {name: index for index, name in enumerate(MONTH_NAMES, start=1)}

COUNTERPARTY_LABEL = "Контрагент"
CONTRACTS_LABEL_DEFAULT = "Договоры"
PROPERTY_LABEL = "По арендованному имуществу"
AMORTIZATION_LABEL = "Амортизация"
TOTAL_LABEL = "Итого"
CLEANED_LABEL = "очищенная"

CURRENCY_FORMAT = '_-* #,##0.00\\ _₽_-;\\-* #,##0.00\\ _₽_-;_-* "-"??\\ _₽_-;_-@_-'

_INVALID_SHEET_CHARS = set(":\\/?*[]")


@dataclass(frozen=True)
class TaxRateRule:
    from_year: int
    multiplier: float


# Tax multipliers are hardcoded here for now; a shared tax-rules tab covering
# every task (including Аренда) is planned, at which point this table will
# stop owning its own copy.
DEFAULT_TAX_RULES: tuple[TaxRateRule, ...] = (
    TaxRateRule(from_year=0, multiplier=0.8),
    TaxRateRule(from_year=2025, multiplier=0.75),
)


def _tax_multiplier_for_year(year: int, tax_rules: Sequence[TaxRateRule] = DEFAULT_TAX_RULES) -> float:
    selected: float | None = None
    for rule in sorted(tax_rules, key=lambda item: item.from_year):
        if year >= rule.from_year:
            selected = rule.multiplier
    if selected is None:
        raise ValueError("tax_rules must contain a rule with from_year <= the requested year")
    return selected


def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    return " ".join(text.split())


@dataclass(frozen=True)
class Period:
    year: int
    month_index: int

    @property
    def month_name(self) -> str:
        return MONTH_NAMES[self.month_index - 1]


@dataclass
class ContractRow:
    contract_number: str
    values: dict[Period, object] = field(default_factory=dict)


@dataclass
class DepreciationTable:
    table_title: str
    contracts_label: str = CONTRACTS_LABEL_DEFAULT
    periods: set[Period] = field(default_factory=set)
    contracts: list[ContractRow] = field(default_factory=list)

    def find_contract(self, contract_number: str) -> ContractRow | None:
        normalized = _normalize_text(contract_number)
        for row in self.contracts:
            if _normalize_text(row.contract_number) == normalized:
                return row
        return None


@dataclass(frozen=True)
class ContractInput:
    contract_number: str
    contract_file: str | Path


@dataclass(frozen=True)
class ApplyResult:
    contract_number: str
    period: Period
    period_created: bool
    contract_created: bool
    value_written: bool


@dataclass(frozen=True)
class GenerationResult:
    workbook: openpyxl.Workbook
    sheet_name: str
    messages: list[str]


def _safe_sheet_title(name: str) -> str:
    cleaned = "".join(ch for ch in _normalize_text(name) if ch not in _INVALID_SHEET_CHARS).strip()
    if not cleaned:
        raise ValueError("Название таблицы не может быть пустым.")
    return cleaned[:31]


def list_table_names(path: str | Path) -> list[str]:
    file_path = Path(path)
    if file_path.suffix.lower() != ".xlsx":
        raise ValueError("Файл итоговой таблицы должен быть в формате .xlsx.")
    workbook = openpyxl.load_workbook(file_path, read_only=True)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def _parse_table_from_sheet(sheet: Worksheet) -> DepreciationTable:
    table_title = _normalize_text(sheet.cell(row=1, column=4).value)
    contracts_label = _normalize_text(sheet.cell(row=2, column=2).value) or CONTRACTS_LABEL_DEFAULT

    period_to_column: dict[Period, int] = {}
    current_month_name: str | None = None
    for column in range(3, sheet.max_column + 1):
        row3_value = _normalize_text(sheet.cell(row=3, column=column).value)
        if row3_value:
            if row3_value == TOTAL_LABEL:
                current_month_name = None
                continue
            current_month_name = row3_value.lower()

        if current_month_name is None or current_month_name not in MONTH_TO_INDEX:
            continue

        year_value = sheet.cell(row=4, column=column).value
        if year_value in (None, ""):
            continue
        try:
            year = int(year_value)
        except (TypeError, ValueError):
            continue

        period_to_column[Period(year=year, month_index=MONTH_TO_INDEX[current_month_name])] = column

    contracts: list[ContractRow] = []
    row = 6
    while row <= sheet.max_row:
        label = _normalize_text(sheet.cell(row=row, column=2).value)
        if not label:
            row += 1
            continue
        if label.lower() == TOTAL_LABEL.lower():
            break

        values: dict[Period, object] = {}
        for period, column in period_to_column.items():
            cell_value = sheet.cell(row=row, column=column).value
            if cell_value not in (None, ""):
                values[period] = cell_value
        contracts.append(ContractRow(contract_number=label, values=values))
        row += 1

    return DepreciationTable(
        table_title=table_title,
        contracts_label=contracts_label,
        periods=set(period_to_column.keys()),
        contracts=contracts,
    )


def load_depreciation_table(path: str | Path, table_name: str) -> DepreciationTable:
    workbook = openpyxl.load_workbook(path, data_only=False)
    try:
        if table_name not in workbook.sheetnames:
            raise ValueError(f"В файле нет таблицы '{table_name}'.")
        return _parse_table_from_sheet(workbook[table_name])
    finally:
        workbook.close()


def create_empty_table(table_title: str) -> DepreciationTable:
    normalized = _normalize_text(table_title)
    if not normalized:
        raise ValueError("Название таблицы не может быть пустым.")
    return DepreciationTable(table_title=normalized)


def apply_contract_period_value(
    table: DepreciationTable,
    contract_number: str,
    period: Period,
    value: float | None,
) -> ApplyResult:
    period_created = period not in table.periods
    table.periods.add(period)

    row = table.find_contract(contract_number)
    contract_created = row is None
    if row is None:
        row = ContractRow(contract_number=_normalize_text(contract_number))
        table.contracts.append(row)

    value_written = False
    if value is not None and period not in row.values:
        row.values[period] = value
        value_written = True

    return ApplyResult(
        contract_number=contract_number,
        period=period,
        period_created=period_created,
        contract_created=contract_created,
        value_written=value_written,
    )


def find_contract_depreciation_total(
    contract_file: str | Path,
    statement_file: str | Path,
    period: Period,
) -> float | None:
    """Find the assets rented under `contract_file` inside `statement_file` for `period`
    and sum their "Начисление амортизации (износа)" / "За период" values.

    Not implemented yet - matching assets listed in a contract registry against rows of
    the Ведомость Амортизации is the next milestone. Returning None leaves the
    corresponding cell blank instead of writing a wrong number.
    """
    return None


def _build_period_grid(periods: Iterable[Period]) -> list[Period]:
    months = sorted({period.month_index for period in periods})
    years = sorted({period.year for period in periods})
    return [Period(year=year, month_index=month_index) for month_index in months for year in years]


def _apply_block_border(sheet: Worksheet, min_row: int, max_row: int, min_col: int, max_col: int, style: str = "thin") -> None:
    side = Side(style=style)
    border = Border(top=side, bottom=side, left=side, right=side)
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            sheet.cell(row=row, column=column).border = border


def _apply_row_border(sheet: Worksheet, row: int, min_col: int, max_col: int, *, top: str | None = None, bottom: str | None = None) -> None:
    top_side = Side(style=top) if top else Side(style="thin")
    bottom_side = Side(style=bottom) if bottom else Side(style="thin")
    thin_side = Side(style="thin")
    for column in range(min_col, max_col + 1):
        sheet.cell(row=row, column=column).border = Border(top=top_side, bottom=bottom_side, left=thin_side, right=thin_side)


def _apply_row_font(sheet: Worksheet, row: int, min_col: int, max_col: int, *, bold: bool = False) -> None:
    for column in range(min_col, max_col + 1):
        sheet.cell(row=row, column=column).font = Font(bold=bold)


def _write_month_headers(sheet: Worksheet, period_grid: Sequence[Period]) -> dict[Period, int]:
    month_groups: list[tuple[str, list[Period]]] = []
    for period in period_grid:
        if month_groups and month_groups[-1][0] == period.month_name:
            month_groups[-1][1].append(period)
        else:
            month_groups.append((period.month_name, [period]))

    period_to_column: dict[Period, int] = {}
    column = 3
    for month_name, month_periods in month_groups:
        start_column = column
        for period in sorted(month_periods, key=lambda item: item.year):
            period_to_column[period] = column
            sheet.cell(row=4, column=column).value = period.year
            sheet.cell(row=5, column=column).value = AMORTIZATION_LABEL
            column += 1
        end_column = column - 1
        sheet.merge_cells(start_row=3, start_column=start_column, end_row=3, end_column=end_column)
        sheet.cell(row=3, column=start_column).value = month_name

    return period_to_column


def _build_year_columns(period_grid: Sequence[Period], start_column: int) -> tuple[dict[int, int], int]:
    years = sorted({period.year for period in period_grid})
    year_to_column: dict[int, int] = {}
    column = start_column
    for year in years:
        year_to_column[year] = column
        column += 1
    return year_to_column, column - 1


def _write_total_headers(sheet: Worksheet, year_to_column: dict[int, int]) -> None:
    if not year_to_column:
        return
    start_column = min(year_to_column.values())
    end_column = max(year_to_column.values())
    sheet.merge_cells(start_row=3, start_column=start_column, end_row=3, end_column=end_column)
    sheet.cell(row=3, column=start_column).value = TOTAL_LABEL
    for year, column in year_to_column.items():
        sheet.cell(row=4, column=column).value = year


def _render_into_sheet(sheet: Worksheet, table: DepreciationTable) -> None:
    if not table.contracts:
        raise ValueError("В таблице нет ни одного договора.")

    period_grid = _build_period_grid(table.periods)
    period_to_column = _write_month_headers(sheet, period_grid)
    last_month_column = max(period_to_column.values(), default=2)
    year_to_column, year_last_column = _build_year_columns(period_grid, last_month_column + 1)
    _write_total_headers(sheet, year_to_column)
    last_column = max(last_month_column, year_last_column)
    last_column_letter = get_column_letter(last_column)

    _apply_block_border(sheet, 1, 5, 1, last_column)

    sheet.merge_cells(f"D1:{last_column_letter}1")
    sheet["D1"] = table.table_title
    sheet["D1"].font = Font(bold=True)
    sheet["D1"].border = Border(top=Side(style="medium"), bottom=Side(style="thin"), left=Side(style="medium"), right=Side(style="medium"))

    sheet.merge_cells("A2:A5")
    sheet.merge_cells("B2:B5")
    sheet.merge_cells(f"C2:{last_column_letter}2")
    sheet["A2"] = COUNTERPARTY_LABEL
    sheet["B2"] = table.contracts_label
    sheet["C2"] = PROPERTY_LABEL

    sheet.column_dimensions["A"].width = 16
    sheet.column_dimensions["B"].width = 22
    for column in range(3, last_column + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 12

    row = 6
    contract_start = row
    for contract in table.contracts:
        sheet.cell(row=row, column=2).value = contract.contract_number

        for period, column in period_to_column.items():
            value = contract.values.get(period)
            if value is not None:
                cell = sheet.cell(row=row, column=column)
                cell.value = value
                cell.number_format = CURRENCY_FORMAT

        for year, year_column in year_to_column.items():
            month_columns = [period_to_column[period] for period in period_grid if period.year == year]
            if not month_columns:
                continue
            terms = "+".join(f"{get_column_letter(column)}{row}" for column in month_columns)
            cell = sheet.cell(row=row, column=year_column)
            cell.value = f"={terms}"
            cell.number_format = CURRENCY_FORMAT

        _apply_row_border(sheet, row, 1, last_column)
        row += 1
    contract_end = row - 1

    total_row = row
    cleaned_row = row + 1
    sheet.cell(row=total_row, column=2).value = TOTAL_LABEL
    sheet.cell(row=cleaned_row, column=2).value = CLEANED_LABEL

    for period, column in list(period_to_column.items()):
        column_letter = get_column_letter(column)
        total_cell = sheet.cell(row=total_row, column=column)
        total_cell.value = f"=SUM({column_letter}{contract_start}:{column_letter}{contract_end})"
        total_cell.number_format = CURRENCY_FORMAT

        multiplier = _tax_multiplier_for_year(period.year)
        cleaned_cell = sheet.cell(row=cleaned_row, column=column)
        cleaned_cell.value = f"={column_letter}{total_row}*{multiplier}"
        cleaned_cell.number_format = CURRENCY_FORMAT

    for year, column in year_to_column.items():
        column_letter = get_column_letter(column)
        total_cell = sheet.cell(row=total_row, column=column)
        total_cell.value = f"=SUM({column_letter}{contract_start}:{column_letter}{contract_end})"
        total_cell.number_format = CURRENCY_FORMAT

        multiplier = _tax_multiplier_for_year(year)
        cleaned_cell = sheet.cell(row=cleaned_row, column=column)
        cleaned_cell.value = f"={column_letter}{total_row}*{multiplier}"
        cleaned_cell.number_format = CURRENCY_FORMAT

    _apply_row_font(sheet, total_row, 1, last_column, bold=True)
    _apply_row_font(sheet, cleaned_row, 1, last_column, bold=True)
    _apply_row_border(sheet, total_row, 1, last_column, top="medium")
    _apply_row_border(sheet, cleaned_row, 1, last_column, bottom="medium")


def _replace_or_create_sheet(workbook: openpyxl.Workbook, sheet_name: str) -> Worksheet:
    safe_name = _safe_sheet_title(sheet_name)
    if safe_name in workbook.sheetnames:
        index = workbook.sheetnames.index(safe_name)
        del workbook[safe_name]
        return workbook.create_sheet(title=safe_name, index=index)
    return workbook.create_sheet(title=safe_name)


def generate_depreciation_workbook(
    *,
    period: Period,
    statement_file: str | Path,
    contracts: Sequence[ContractInput],
    source_table_path: str | Path | None = None,
    source_table_name: str | None = None,
    new_table_name: str | None = None,
) -> GenerationResult:
    if not contracts:
        raise ValueError("Добавьте хотя бы один договор аренды.")

    if source_table_path is not None:
        workbook = openpyxl.load_workbook(source_table_path, data_only=False)
    else:
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

    messages: list[str] = []

    if source_table_name:
        if source_table_name not in workbook.sheetnames:
            raise ValueError(f"В файле нет таблицы '{source_table_name}'.")
        table = _parse_table_from_sheet(workbook[source_table_name])
        target_sheet_name = source_table_name
        messages.append(f"Дополняется существующая таблица '{source_table_name}'.")
    else:
        if not new_table_name or not new_table_name.strip():
            raise ValueError("Укажите название новой таблицы.")
        if source_table_path is not None and new_table_name.strip() in workbook.sheetnames:
            raise ValueError(f"В файле уже есть таблица '{new_table_name.strip()}'.")
        table = create_empty_table(new_table_name.strip())
        target_sheet_name = table.table_title
        messages.append(f"Создана новая таблица '{table.table_title}'.")

    period_already_existed = period in table.periods
    contract_messages: list[str] = []
    for contract in contracts:
        value = find_contract_depreciation_total(contract.contract_file, statement_file, period)
        result = apply_contract_period_value(table, contract.contract_number, period, value)
        if result.contract_created:
            contract_messages.append(f"Договор '{contract.contract_number}': добавлена новая строка в таблице.")
        elif result.value_written:
            contract_messages.append(f"Договор '{contract.contract_number}': записано новое значение.")
        else:
            contract_messages.append(f"Договор '{contract.contract_number}': строка уже существовала, значение не изменено.")

    if period_already_existed:
        messages.append(f"Месяц {period.month_name} {period.year} уже существовал в таблице — новый столбец не добавлялся.")
    else:
        messages.append(f"Добавлен новый столбец периода: {period.month_name} {period.year}.")
    messages.extend(contract_messages)

    sheet = _replace_or_create_sheet(workbook, target_sheet_name)
    _render_into_sheet(sheet, table)
    workbook.active = workbook.sheetnames.index(sheet.title)

    return GenerationResult(workbook=workbook, sheet_name=sheet.title, messages=messages)


def save_depreciation_workbook(result: GenerationResult, output_path: str | Path) -> None:
    result.workbook.save(output_path)
