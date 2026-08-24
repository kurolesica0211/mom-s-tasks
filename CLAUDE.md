# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Tkinter desktop app ("Инструменты бухгалтера" / Accountant Tools) that generates Excel reports for a Russian accounting practice. Each "task" in the app reads one or more source Excel exports (legacy `.xls` accounting-system exports, or `.xlsx`) and produces a new formatted workbook. It's distributed as a standalone executable (PyInstaller), not run as a normal Python package by end users.

## Commands

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python3 depreciation_ui.py          # run the app (this is the entry point, despite the name)

# Build the standalone executable (matches .github/workflows/buid.yml, triggered manually from Actions)
pyinstaller --onefile --windowed --name AccountantTools depreciation_ui.py
```

There is no test suite, linter, or formatter configured in this repo.

## Architecture

**UI/logic split.** Each task's Excel-generation logic lives in its own `<task>_table.py` module — pure functions and dataclasses, using `openpyxl`/`xlrd` directly, with zero Tkinter imports. All UI code lives in `depreciation_ui.py` in a single class (`ArendaUI` — the name predates the second task and is now a misnomer for the whole app). This split is deliberate and should be preserved for any new task: build `<task>_table.py` first, then wire it into the UI.

- `arenda_table.py` — Аренда (rent) task: parses "Анализ счёта 90.01" exports, aggregates rent by lessor/contract/period.
- `depreciation_table.py` — Амортизация (depreciation) task: the more involved one. It reads a "Ведомость Амортизации" statement plus one or more "Договор" (rental contract) registries, and for each contract must locate the right assets in the statement to sum their period depreciation charge. That matching step has real fuzziness (see `match_contract_depreciation` and friends): it tries an "Инв. №" exact match against the statement's inventory-number column first, falls back to a substring search of a VIN/factory-number across every column, splits compound values, and normalizes Cyrillic/Latin look-alike characters (e.g. Cyrillic `Х`/`З` vs Latin `X`/`3`) since real registries mix them. Ambiguous or missing matches are collected as `AssetFailure`s and surfaced to the user rather than guessed at.
- `tax_rules.py` — the global "tax multiplier" rules (`TaxRateRule`, `DEFAULT_TAX_RULES`, `tax_multiplier_for_year`). Any task that needs an after-tax value (an "очищенная" column, etc.) should import from here rather than hardcoding multipliers. The UI exposes one shared, editable rules window (button in the header, reachable from every task) backed by this module — there is intentionally no per-task tax config anymore.

**Column widths and number formats are auto-fit, never hardcoded.** Both `<task>_table.py` files size every column to its actual rendered content via `_autosize_columns`/`_cell_text_width`/`_wide_merge_anchors` (duplicated in each file, same as `CURRENCY_FORMAT` — task modules don't cross-import each other). This matters because the app's two most common viewers disagree about unset widths: Apple Numbers auto-fits on import regardless of what's stored, while Excel (Windows or Mac) renders the stored width literally — so every column needs an explicit, correct width or Windows/Excel users see cramped columns that Mac/Numbers users don't. Since openpyxl never evaluates formulas, a formula cell (`=SUM(...)`, `=B6+C6`, etc.) can't be measured directly from its cell value — the rendering functions compute the real total in Python alongside building the formula string and feed it through an `extra_widths`/`column_extra` dict. Any new formula-writing code must do the same or its column silently falls back to `min_width`. Relatedly: `CURRENCY_FORMAT`'s code must stay in invariant syntax (`,` is always the thousands placeholder, `.` is always the decimal placeholder — Excel translates those to the viewer's own locale glyphs at render time, the file never controls that). Writing them swapped, as `arenda_table.py` once did (`'#.##0,00'`), produces a malformed format code that different viewers handle inconsistently instead of a "Russian-style" display.

**Task panel pattern in `depreciation_ui.py`.** The header holds a task dropdown; each task gets a panel built once at startup (`_build_<task>_panel`) and shown/hidden via `grid()`/`grid_forget()` in `_on_task_changed` rather than rebuilt each time. Panels with scrollable lists (lessors, contract files) use a repeated canvas+scrollbar+inner-frame pattern with manual mousewheel binding (see `_on_*_mousewheel`) — copy that pattern for new scrollable sections rather than inventing another one.

**File-handling convention: never mutate the source file.** Every task always writes to a brand-new file chosen via a save dialog, even when the user is "extending" an existing table. `depreciation_table.py`'s approach is the model to follow: fully parse the existing sheet into an in-memory dataclass model, apply the requested changes to that model, then rebuild the sheet from scratch (preserving any other sheets in the source workbook untouched). This avoids ever hand-editing merged cells/formulas in place.

**Indentation is not uniform across files** — `arenda_table.py` uses tabs; `depreciation_ui.py`, `depreciation_table.py`, and `tax_rules.py` use 4-space indentation. Match whichever file you're editing.

**`example_tables/`** holds real (example) accounting export files used for manual validation against known-correct values in existing workbooks — there's no pytest fixture harness, validation is done by running the actual matching functions against these files and comparing to figures already present in a real output workbook.

**`fee_returns_task.py`** is an unrelated one-off pandas script (extracts commission/return amounts from a specific bank statement filename) — not part of the Tkinter app, not imported by anything else.

## Keeping this file current

Update this file whenever something changes that would otherwise cost a fresh session real time to rediscover — a new task module, a changed convention, a resolved gotcha, a new shared module. Keep entries short and specific: this file is meant to be read in full at the start of a session, not searched. Prefer editing or removing a stale bullet over leaving it and appending a contradicting one next to it.
