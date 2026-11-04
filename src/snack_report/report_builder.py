"""Сборка ежедневного отчёта о наличии снеков и напитков (.xlsx)."""
from __future__ import annotations

import datetime as dt
import re
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

REPORT_PREFIX = "Отчёт_"
REPORT_NAME_RE = re.compile(r"^Отчёт_(\d{2}\.\d{2}\.\d{4})\.xlsx$")

# Названия колонок в строке заголовка отчёта
COL_INCOME = "Приход"
COL_MORNING = "ОстУтро"
COL_SOLD = "Продано"
COL_DEBT = "Долг"
COL_WHO = "Имя"
COL_WRITEOFF = "Списание"
COL_REASON = "Причина"
COL_EVENING = "ОстВечер"
CLEARED_COLS = (COL_INCOME, COL_SOLD, COL_DEBT, COL_WHO, COL_WRITEOFF, COL_REASON)

# Не товары. Разделы копируются как есть, строки персонала очищаются каждый день.
DEFAULT_SECTION_ROWS = ("Прочее",)
DEFAULT_STAFF_ROWS = ("Утеря карт", "Наличие карт", "Администратор", "Тех персонал", "Охрана")


class ReportFormatError(Exception):
    pass


def normalize(name: str) -> str:
    """Для сравнения названий: без учёта регистра, ё/е, запятой/точки и лишних пробелов."""
    s = str(name).lower().replace("ё", "е").replace(",", ".")
    return re.sub(r"\s+", " ", s).strip()


def report_filename(date: dt.date) -> str:
    return f"{REPORT_PREFIX}{date:%d.%m.%Y}.xlsx"


def find_previous_report(reports_dir: Path, before: dt.date) -> Path | None:
    """Самый свежий Отчёт_ДД.ММ.ГГГГ.xlsx с датой раньше `before`."""
    best: tuple[dt.date, Path] | None = None
    if not reports_dir.is_dir():
        return None
    for p in reports_dir.iterdir():
        m = REPORT_NAME_RE.match(p.name)
        if not m:
            continue
        d = dt.datetime.strptime(m.group(1), "%d.%m.%Y").date()
        if d < before and (best is None or d > best[0]):
            best = (d, p)
    return best[1] if best else None


@dataclass
class Layout:
    header_row: int
    cols: dict[str, int]          # название колонки -> номер (1-based)
    item_rows: dict[str, int]     # нормализованное название товара -> строка
    item_names: dict[str, str]    # нормализованное -> как написано в отчёте
    staff_rows: list[int]
    first_section_row: int | None  # первая строка-раздел/персонал после товаров основного блока


NAME_COL = 1


def detect_layout(ws: Worksheet, section_rows: tuple[str, ...], staff_rows: tuple[str, ...]) -> Layout:
    for r in range(1, min(ws.max_row, 15) + 1):
        values = {str(c.value).strip(): c.column for c in ws[r] if c.value is not None}
        if COL_MORNING in values and COL_EVENING in values:
            header_row, cols = r, values
            break
    else:
        raise ReportFormatError(f"В шаблоне не найдена строка заголовка с «{COL_MORNING}» и «{COL_EVENING}»")

    missing = [c for c in (COL_INCOME, COL_MORNING, COL_SOLD, COL_WRITEOFF, COL_EVENING) if c not in cols]
    if missing:
        raise ReportFormatError(f"В заголовке отчёта нет колонок: {', '.join(missing)}")

    sections = {normalize(s) for s in section_rows}
    staff = {normalize(s) for s in staff_rows}
    lay = Layout(header_row, cols, {}, {}, [], None)
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, NAME_COL).value
        if v is None or not str(v).strip():
            continue
        key = normalize(v)
        if key in sections or key in staff:
            if lay.first_section_row is None:
                lay.first_section_row = r
            if key in staff:
                lay.staff_rows.append(r)
            continue
        if key in lay.item_rows:
            raise ReportFormatError(f"Товар «{v}» встречается в отчёте дважды (строки {lay.item_rows[key]} и {r})")
        lay.item_rows[key] = r
        lay.item_names[key] = str(v).strip()
    return lay


def _num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ".").strip())
    except ValueError:
        return 0.0


def _clean(x: float):
    return int(x) if float(x).is_integer() else round(x, 3)


def evening_value(ws: Worksheet, ws_values: Worksheet, cols: dict[str, int], r: int) -> float:
    """ОстВечер строки: значение формулы, если файл сохранялся в Excel, иначе считаем сами."""
    cached = ws_values.cell(r, cols[COL_EVENING]).value
    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
        return float(cached)
    g = lambda col: _num(ws.cell(r, cols[col]).value)  # noqa: E731
    return g(COL_MORNING) + g(COL_INCOME) - g(COL_SOLD) - g(COL_WRITEOFF)


def _insert_items(ws: Worksheet, lay: Layout, names: list[str]) -> None:
    """Вставляет новые товары в конец основного блока (перед «Прочее»), копируя оформление строки выше."""
    at = lay.first_section_row or ws.max_row + 1
    ws.insert_rows(at, len(names))
    style_row = at - 1
    for i, name in enumerate(names):
        r = at + i
        for c in range(1, ws.max_column + 1):
            src = ws.cell(style_row, c)
            dst = ws.cell(r, c)
            if src.has_style:
                dst._style = copy(src._style)
        ws.cell(r, NAME_COL).value = name
        if ws.row_dimensions[style_row].height:
            ws.row_dimensions[r].height = ws.row_dimensions[style_row].height


def _reset_negative_highlight(ws: Worksheet, lay: Layout) -> None:
    """insert_rows не сдвигает условное форматирование — задаём его заново."""
    ws.conditional_formatting = ConditionalFormattingList()
    col = get_column_letter(lay.cols[COL_EVENING])
    ws.conditional_formatting.add(
        f"{col}{lay.header_row + 1}:{col}{ws.max_row}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006")),
    )


@dataclass
class BuildResult:
    output: Path
    source: Path
    from_template: bool
    filled: list[tuple[str, float]] = field(default_factory=list)
    added: list[str] = field(default_factory=list)          # новые строки (названия из UCS как есть)
    negative: list[tuple[str, float]] = field(default_factory=list)


def build_report(
    *,
    date: dt.date,
    sales: dict[str, float],
    source: Path,
    from_template: bool,
    output: Path,
    section_rows: tuple[str, ...] = DEFAULT_SECTION_ROWS,
    staff_rows: tuple[str, ...] = DEFAULT_STAFF_ROWS,
) -> BuildResult:
    """
    source — вчерашний отчёт (ОстУтро = вчерашний ОстВечер)
             или шаблон (ОстУтро берётся как есть, его вносят вручную при первом запуске).
    Названия из UCS переносятся как есть: продажа идёт в строку с тем же названием,
    а если такой строки нет — добавляется новая.
    """
    wb = openpyxl.load_workbook(source)
    wb_values = openpyxl.load_workbook(source, data_only=True)  # значения формул, если файл сохранялся в Excel
    ws = wb.worksheets[0]
    ws_values = wb_values.worksheets[0]
    lay = detect_layout(ws, section_rows, staff_rows)
    c = lay.cols

    # 1. Утренние остатки (по вчерашнему файлу, до вставки новых строк)
    morning: dict[str, float] = {}
    for key, r in lay.item_rows.items():
        if from_template:
            morning[key] = _num(ws.cell(r, c[COL_MORNING]).value)
        else:
            morning[key] = evening_value(ws, ws_values, c, r)

    # 2. Продажи: суммируем по нормализованному названию, новые товары добавляем строками
    sold: dict[str, float] = {}
    new_names: dict[str, str] = {}
    for name, qty in sales.items():
        key = normalize(name)
        sold[key] = sold.get(key, 0) + qty
        if key not in lay.item_rows and key not in new_names:
            new_names[key] = name
    result = BuildResult(output=output, source=source, from_template=from_template, added=list(new_names.values()))
    if new_names:
        _insert_items(ws, lay, list(new_names.values()))
        lay = detect_layout(ws, section_rows, staff_rows)
        c = lay.cols

    # 3. Заполнение
    date_cell = ws.cell(max(lay.header_row - 1, 1), NAME_COL)
    date_cell.value = date
    date_cell.number_format = "DD.MM.YYYY"

    L = lambda col: get_column_letter(c[col])  # noqa: E731
    for key, r in lay.item_rows.items():
        for col in CLEARED_COLS:
            if col in c:
                ws.cell(r, c[col]).value = None
        m = morning.get(key, 0.0)
        ws.cell(r, c[COL_MORNING]).value = _clean(m)
        s = sold.get(key)
        if s:
            ws.cell(r, c[COL_SOLD]).value = _clean(s)
            result.filled.append((lay.item_names[key], s))
        ws.cell(r, c[COL_EVENING]).value = (
            f"={L(COL_MORNING)}{r}+{L(COL_INCOME)}{r}-{L(COL_SOLD)}{r}-{L(COL_WRITEOFF)}{r}"
        )
        if m - (s or 0) < 0:
            result.negative.append((lay.item_names[key], m - (s or 0)))

    # 4. Персонал и карты — каждый день с чистого листа
    for r in lay.staff_rows:
        for col in range(NAME_COL + 1, ws.max_column + 1):
            cell = ws.cell(r, col)
            if not isinstance(cell, MergedCell):
                cell.value = None

    _reset_negative_highlight(ws, lay)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return result
