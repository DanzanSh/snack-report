"""Ежемесячный отчёт: итоги по дневным отчётам Отчёт_ДД.ММ.ГГГГ.xlsx за месяц."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .cash_report import get_snack_sheet
from .report_builder import (
    COL_DEBT,
    COL_INCOME,
    COL_MORNING,
    COL_SOLD,
    COL_WRITEOFF,
    DEFAULT_SECTION_ROWS,
    DEFAULT_STAFF_ROWS,
    REPORT_NAME_RE,
    _clean,
    _num,
    detect_layout,
    evening_value,
)

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
HEADERS = ["Товар", "ОстНачало", "Приход", "Продано", "Долг", "Списание", "ОстКонец"]
SUM_FIELDS = ("income", "sold", "debt", "writeoff")


def monthly_filename(year: int, month: int) -> str:
    return f"Месяц_{month:02d}.{year}.xlsx"


@dataclass
class DayRow:
    morning: float
    income: float
    sold: float
    debt: float
    writeoff: float
    evening: float


def read_day(path: Path, section_rows, staff_rows) -> list[tuple[str, DayRow]]:
    """Товары дневного отчёта в порядке строк."""
    ws = get_snack_sheet(openpyxl.load_workbook(path))
    ws_values = get_snack_sheet(openpyxl.load_workbook(path, data_only=True))
    lay = detect_layout(ws, section_rows, staff_rows)
    c = lay.cols
    g = lambda r, col: _num(ws.cell(r, c[col]).value) if col in c else 0.0  # noqa: E731
    rows = []
    for key, r in sorted(lay.item_rows.items(), key=lambda kv: kv[1]):
        rows.append((lay.item_names[key], DayRow(
            morning=g(r, COL_MORNING), income=g(r, COL_INCOME), sold=g(r, COL_SOLD),
            debt=g(r, COL_DEBT), writeoff=g(r, COL_WRITEOFF), evening=evening_value(ws, ws_values, c, r),
        )))
    return rows


@dataclass
class Item:
    start: float
    end: float
    income: float = 0.0
    sold: float = 0.0
    debt: float = 0.0
    writeoff: float = 0.0


@dataclass
class MonthResult:
    output: Path
    days: int


def daily_reports(folder: Path) -> dict[dt.date, Path]:
    found = {}
    for p in folder.iterdir() if folder.is_dir() else []:
        m = REPORT_NAME_RE.match(p.name)
        if m:
            found[dt.datetime.strptime(m.group(1), "%d.%m.%Y").date()] = p
    return found


def build_month(
    year: int,
    month: int,
    reports: dict[dt.date, Path],
    output: Path,
    section_rows=DEFAULT_SECTION_ROWS,
    staff_rows=DEFAULT_STAFF_ROWS,
) -> MonthResult:
    """reports — дневные отчёты этого месяца. Дни без отчёта (выходные, праздники) не считаются пропусками."""
    items: dict[str, Item] = {}   # порядок вставки = порядок первого появления
    last_names: list[str] = []
    for d in sorted(reports):
        day_rows = read_day(reports[d], section_rows, staff_rows)
        for name, day in day_rows:
            it = items.get(name)
            if it is None:  # ОстНачало — ОстУтро первого дня, когда товар есть в отчёте
                items[name] = it = Item(start=day.morning, end=day.evening)
            it.end = day.evening
            for f in SUM_FIELDS:
                setattr(it, f, getattr(it, f) + getattr(day, f))
        last_names = [n for n, _ in day_rows]
    # порядок как в последнем дневном отчёте, исчезнувшие из него товары — в конце
    order = last_names + [n for n in items if n not in last_names]

    _write(output, year, month, order, items, sorted(reports))
    return MonthResult(output=output, days=len(reports))


def _write(output: Path, year: int, month: int, order, items, dates) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{MONTHS[month - 1].capitalize()} {year}"
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    beige = PatternFill("solid", fgColor="F2DCCB")
    ncol = len(HEADERS)
    last_col = get_column_letter(ncol)

    ws["A1"] = f"Отчёт о наличии снеков и напитков за {MONTHS[month - 1]} {year}"
    ws["A1"].font = Font(bold=True, size=12)
    ws.merge_cells(f"A1:{last_col}1")
    ws["A2"] = f"Дневных отчётов: {len(dates)}, с {dates[0]:%d.%m.%Y} по {dates[-1]:%d.%m.%Y}"
    ws.merge_cells(f"A2:{last_col}2")

    header_row = 4
    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(header_row, i, h)
        cell.font = Font(bold=True)
        cell.fill = beige
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    r = header_row
    for name in order:
        it = items[name]
        r += 1
        values = [name, it.start, it.income, it.sold, it.debt, it.writeoff, it.end]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(r, i, v if i == 1 else _clean(v))
            cell.border = border
            if i > 1:
                cell.alignment = Alignment(horizontal="center")
    first, last = header_row + 1, r

    total = r + 1
    ws.cell(total, 1, "ИТОГО").font = Font(bold=True)
    for i in range(2, ncol + 1):
        col = get_column_letter(i)
        cell = ws.cell(total, i, f"=SUM({col}{first}:{col}{last})")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for i in range(1, ncol + 1):
        ws.cell(total, i).border = border
        ws.cell(total, i).fill = beige

    ws.conditional_formatting.add(
        f"{last_col}{first}:{last_col}{last}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006")),
    )
    for i, w in enumerate([32, 11, 10, 10, 8, 10, 11], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(header_row + 1, 2)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def update_monthly(folder: Path, section_rows=DEFAULT_SECTION_ROWS, staff_rows=DEFAULT_STAFF_ROWS) -> list[MonthResult]:
    """Пересобирает месячные отчёты, которых нет или которые старше своих дневных отчётов."""
    reports = daily_reports(folder)
    if not reports:
        return []
    by_month: dict[tuple[int, int], dict[dt.date, Path]] = {}
    for d, p in reports.items():
        by_month.setdefault((d.year, d.month), {})[d] = p

    results = []
    for (y, m), month_reports in sorted(by_month.items()):
        out = folder / monthly_filename(y, m)
        newest_daily = max(p.stat().st_mtime for p in month_reports.values())
        if out.exists() and out.stat().st_mtime >= newest_daily:
            continue
        results.append(build_month(y, m, month_reports, out, section_rows, staff_rows))
    return results
