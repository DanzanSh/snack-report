"""Первый лист дневного отчёта — «Ежедневный кассовый отчет» — и работа с листами книги.

В книге два листа: кассовый (первый) и отчёт о наличии снеков и напитков (второй).
Название листа в Excel не может быть длиннее 31 символа, поэтому полное название
«Ежедневный отчет о наличии снеков и напитков» стоит в заголовке самого листа.
"""
from __future__ import annotations

import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

CASH_SHEET = "Ежедневный кассовый отчет"
SNACK_SHEET = "Наличие снеков и напитков"
CASH_TITLE = "Ежедневный кассовый отчет"
ROAD_EXPENSES = 1000          # «Дорожные» — всегда 1000
DATE_FORMAT = r"dd\.mm\.yyyy"  # как в эталонном отчёте

# Ячейки кассового листа
DATE_CELL = "B3"
ADMIN_CELL = "D3"
OPENING_CELL = "C5"           # остаток наличных на начало дня
ROAD_ROW = 7
EXPENSE_ROWS = range(8, 15)   # прочие расходы: название в B, сумма в C, расшифровка в D
EXPENSE_FIRST, EXPENSE_LAST = 7, 14
CASH_IN_CELL, NONCASH_CELL = "C17", "C18"
INCOME_CELLS = ("C20", "C21", "C22")   # приход, предоплаты, инкасация
CLOSING_CELL = "C23"          # остаток наличных на конец дня
STAFF_CELLS = ("C25", "C26")  # Тех персонал, Охрана

# Подписи и формулы, которые есть в отчёте всегда
LABELS = {
    "B2": CASH_TITLE, "D3": "Администратор ",
    "B4": "Остаток наличных на начала дня", "C4": "Сумма", "B5": "наличные",
    "B6": "Расходы", "D6": "Расшифровка", "B7": "Дорожные",
    "B15": "общая сумма расходов", "B16": "КАССА ЗА ДЕНЬ", "B17": "Наличные", "B18": "Безнал",
    "B19": "Общая выручка", "B20": "приход", "B21": "Предоплаты", "B22": "Инкасация",
    "B23": "Остаток наличных на конец дня", "B24": "На смене были", "D24": "Подпись Администратора",
    "B25": "Тех персонал", "B26": "Охрана",
}
FORMULAS = {
    "C15": f"=SUM(C{EXPENSE_FIRST}:C{EXPENSE_LAST})",
    "C19": "=C17+C18",
    CLOSING_CELL: "=C5+C17+C20+C21-C15",
}
MERGES = ("B2:D2", "D17:D19", "D25:D26")
ROW_HEIGHTS = {1: 16.5, 2: 19.5, 3: 18.6, 4: 18.6, 5: 24.0, 15: 22.15, 16: 22.15, 17: 15.75,
               19: 16.5, 20: 14.25, 22: 16.5, 23: 16.5, 25: 22.5, 26: 27.0}
COL_WIDTHS = {"A": 8.9, "B": 33.4, "C": 14.7, "D": 52.6, "E": 8.9}

HEAD, SUB = "FBE5D6", "F8CBAD"   # заливки
# ячейка, границы (слева, справа, сверху, снизу: m — средняя, t — тонкая, - — нет),
# заливка, размер шрифта, выравнивание по горизонтали и вертикали, перенос, формат числа
_STYLE = [
    ("B2", "mtmm", None, 14, "center", None, 0, ""), ("C2", "--mm", None, 11, None, None, 0, ""),
    ("D2", "-tmm", None, 11, None, None, 0, ""),
    ("B3", "mt-t", None, 12, "center", "center", 0, DATE_FORMAT), ("C3", "tt-t", None, 12, None, None, 0, ""),
    ("D3", "tm-t", None, 12, None, None, 0, ""),
    ("B4", "mttt", HEAD, 12, None, "center", 0, ""), ("C4", "tttt", HEAD, 12, "center", "center", 0, ""),
    ("D4", "tmtt", HEAD, 12, None, None, 0, ""),
    ("B5", "mttt", None, 12, "center", "center", 0, ""), ("C5", "tttt", HEAD, 12, "center", "center", 0, "#,##0"),
    ("D5", "tmtt", None, 12, None, None, 0, ""),
    ("B6", "mttt", HEAD, 12, "center", "center", 0, ""), ("C6", "tttt", HEAD, 12, "center", "center", 0, "#,##0"),
    ("D6", "tmtt", HEAD, 12, "center", None, 0, ""),
    ("B7", "mttt", None, 12, "center", "center", 0, ""), ("C7", "tttt", None, 12, "center", "center", 0, "#,##0"),
    ("D7", "tmtt", None, 12, None, None, 0, ""),
    *[(f"{c}{r}", b, None, 12, h, v, 0, nf)
      for r in range(8, 15)
      for c, b, h, v, nf in (("B", "mttt", "center", "center", ""),
                             ("C", "tttt", "center", "center", "#,##0"),
                             ("D", "tmtt", None, None, ""))],
    ("B15", "mtt-", None, 12, "center", "center", 0, ""), ("C15", "ttt-", None, 12, "center", "center", 0, "#,##0"),
    ("D15", "tmt-", None, 12, None, None, 0, ""),
    ("B16", "mtmm", HEAD, 12, "center", "center", 0, ""), ("C16", "mtmm", HEAD, 12, "center", "center", 0, "#,##0"),
    ("D16", "mmmm", HEAD, 12, "center", "center", 0, ""),
    ("B17", "mt-t", None, 12, None, None, 0, ""), ("C17", "tt-t", None, 12, "center", "center", 0, "#,##0"),
    ("D17", "tmmt", None, 10, "center", None, 1, ""),
    ("B18", "mttt", None, 12, None, None, 0, ""), ("C18", "tttt", None, 12, "center", "center", 0, "#,##0"),
    ("D18", "tm--", None, 11, None, None, 0, ""),
    ("B19", "mttt", None, 12, None, None, 0, ""), ("C19", "tttt", None, 12, "center", "center", 0, "#,##0"),
    ("D19", "tm-t", None, 11, None, None, 0, ""),
    ("B20", "mttt", None, 12, None, None, 0, ""), ("C20", "tttt", None, 12, "center", "center", 0, "#,##0"),
    ("D20", "tmtt", None, 12, "center", "center", 0, ""),
    ("B21", "mttt", None, 12, None, None, 0, ""), ("C21", "tttt", None, 12, "center", "center", 0, "#,##0"),
    ("D21", "tmtt", None, 12, None, None, 0, ""),
    ("B22", "mtt-", None, 12, None, None, 0, ""), ("C22", "ttt-", None, 12, "center", "center", 0, "#,##0"),
    ("D22", "tmt-", None, 12, None, None, 0, ""),
    ("B23", "mmm-", HEAD, 12, "center", "center", 0, ""), ("C23", "-tm-", None, 12, "center", "center", 0, "#,##0"),
    ("D23", "tmm-", None, 12, None, None, 0, ""),
    ("B24", "mtmt", SUB, 12, None, None, 0, ""), ("C24", "ttmt", SUB, 12, None, None, 0, ""),
    ("D24", "tmmt", SUB, 12, "right", None, 0, ""),
    ("B25", "mttt", None, 12, None, None, 0, ""), ("C25", "tttt", None, 12, "center", "center", 0, ""),
    ("D25", "tmtm", None, 12, None, None, 0, ""),
    ("B26", "mttm", None, 12, None, None, 0, ""), ("C26", "tttm", None, 12, "center", "center", 0, ""),
    ("D26", "tm-m", None, 11, None, None, 0, ""),
]
_SIDES = {"m": Side(style="medium"), "t": Side(style="thin"), "-": Side()}


def build_cash_sheet(wb: Workbook, index: int = 0) -> Worksheet:
    """Создаёт пустой кассовый лист (первым): подписи, формулы, «Дорожные» = 1000."""
    ws = wb.create_sheet(CASH_SHEET, index)
    for coord, borders, fill, size, h, v, wrap, nf in _STYLE:
        cell = ws[coord]
        left, right, top, bottom = (_SIDES[b] for b in borders)
        cell.border = Border(left=left, right=right, top=top, bottom=bottom)
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = Font(name="Calibri", size=size)
        cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=bool(wrap))
        if nf:
            cell.number_format = nf
    for coord, text in LABELS.items():
        ws[coord] = text
    ws[f"C{ROAD_ROW}"] = ROAD_EXPENSES
    for coord, formula in FORMULAS.items():
        ws[coord] = formula
    for rng in MERGES:
        ws.merge_cells(rng)
    for r, h in ROW_HEIGHTS.items():
        ws.row_dimensions[r].height = h
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    return ws


def get_cash_sheet(wb: Workbook) -> Worksheet | None:
    return wb[CASH_SHEET] if CASH_SHEET in wb.sheetnames else None


def get_snack_sheet(wb: Workbook) -> Worksheet:
    """Лист «Наличие снеков и напитков»; в старых отчётах (один лист) — первый лист."""
    if SNACK_SHEET in wb.sheetnames:
        return wb[SNACK_SHEET]
    return next(ws for ws in wb.worksheets if ws.title != CASH_SHEET)


def ensure_sheets(wb: Workbook) -> tuple[Worksheet, Worksheet]:
    """Приводит книгу к виду: 1-й лист — кассовый, 2-й — снеки и напитки. Возвращает (касса, снеки)."""
    snack = get_snack_sheet(wb)
    if snack.title != SNACK_SHEET:
        snack.title = SNACK_SHEET
    cash = get_cash_sheet(wb) or build_cash_sheet(wb)
    wb.move_sheet(cash, offset=-wb.index(cash))
    wb.active = 0
    for ws in wb.worksheets:   # выделен только первый лист, иначе Excel откроет книгу «группой» листов
        ws.sheet_view.tabSelected = ws is cash
    return cash, snack


def _num(v) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def closing_cash(ws: Worksheet, ws_values: Worksheet | None) -> float:
    """Остаток наличных на конец дня: значение формулы, если файл сохранялся в Excel, иначе считаем сами."""
    cached = ws_values[CLOSING_CELL].value if ws_values is not None else None
    if isinstance(cached, (int, float)) and not isinstance(cached, bool):
        return float(cached)
    expenses = sum(_num(ws.cell(r, 3).value) for r in range(EXPENSE_FIRST, EXPENSE_LAST + 1))
    return (_num(ws[OPENING_CELL].value) + _num(ws[CASH_IN_CELL].value)
            + _num(ws["C20"].value) + _num(ws["C21"].value) - expenses)


def fill_cash(ws: Worksheet, date: dt.date, opening: float | None) -> None:
    """Готовит кассовый лист на новый день.

    Остаются: дата, «Дорожные» = 1000, подписи и формулы. Очищаются: прочие расходы и их расшифровка,
    выручка, приход, предоплаты, инкасация, администратор и имена персонала.
    opening — остаток на начало дня из прошлого отчёта; None — оставить ячейку как есть (вносят вручную).
    """
    ws[DATE_CELL].value = date
    ws[DATE_CELL].number_format = DATE_FORMAT
    if opening is not None:
        ws[OPENING_CELL].value = int(opening) if float(opening).is_integer() else round(opening, 2)
    for r in EXPENSE_ROWS:
        for col in (2, 3):
            ws.cell(r, col).value = None
    for r in range(ROAD_ROW, EXPENSE_LAST + 1):
        ws.cell(r, 4).value = None
    ws[f"B{ROAD_ROW}"].value = LABELS[f"B{ROAD_ROW}"]
    ws[f"C{ROAD_ROW}"].value = ROAD_EXPENSES
    for coord in (CASH_IN_CELL, NONCASH_CELL, *INCOME_CELLS, "D17", *STAFF_CELLS):
        ws[coord].value = None
    ws[ADMIN_CELL].value = LABELS[ADMIN_CELL]
    for coord, formula in FORMULAS.items():
        ws[coord].value = formula
