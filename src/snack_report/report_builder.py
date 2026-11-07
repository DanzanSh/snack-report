"""Сборка ежедневного отчёта о наличии снеков и напитков (.xlsx)."""
from __future__ import annotations

import datetime as dt
import re
import unicodedata
from fnmatch import fnmatchcase
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

from .cash_report import (
    OPENING_CELL,
    closing_cash,
    ensure_sheets,
    fill_cash,
    get_cash_sheet,
    get_snack_sheet,
)

REPORT_PREFIX = "Отчёт_"
# Отчёт_ГГГГ-ММ-ДД.xlsx; старый формат Отчёт_ДД.ММ.ГГГГ.xlsx тоже читается
REPORT_NAME_RE = re.compile(r"^Отчёт_(\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})\.xlsx$")

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

# Название в UCS -> строка отчёта, когда они пишутся по-разному. Дополняется секцией [aliases] в settings.ini.
DEFAULT_ALIASES = {
    "Протеин 30гр": "Протеин в пачках штучно",
    "протеиновый коктейль (на молоке)": "Протеин порц",
    "протеиновый коктейль на воде": "Протеин порц",
    "конфета 20 грамм": "Конфеты штучно",
    "RedBull 0.25": "Ред Булл",
    "Бомбар Витамин и минерал": "Энергетик Бомбар Вит мин",
    "Вода 1 л": "Вода 1",
    "Вода 1,5 л": "Вода 1,5",
    "энергетик вейк апп": "энергетик вейкапп",
    "утеря карточки": "Утеря карт",
    # снеки и еда
    "Зефир": "Зефир 750",
    "Зефир Большой": "Зефир 1000",
    "Козинак": "Козинаки",
    "Козинак мини": "Козинаки",
    "Печенье 0 кал": "печенье барное",
    "Сникерс Торт": "торт Сникерс",
    "Рис яйца": "Рис с яйцами",
    # вещи
    "Полотенце": "Полотенца",
    "Резинки": "резинка для волос",
    "Тапочки": "тапочки одноразовые",
    "Футболка 777 Клиенты": "Футболки 777",
    "Футболка для тренеров": "Футболки 777",
    # напитки и порции
    "Американо с молоком": "Американо",
    "BCAA": "BCAA порц",
    "Гейнер на воде": "Гейнер",
    "Гейнер на Молоке": "Гейнер",
    "каллоген банка 300грамм": "калоген?",
    "каллоген порционный": "каллоген порция",
    "Креатин порция": "Креатин порц",
    "соки грацио стекло": "Сок Грацио стекло 0.2",
    # банки и упаковки
    "Now Omega 3 Gold 200 Caps": "Омега",
    "Psihotic 220gramm": "Психотик",
    "Креатин 300 грамм": "Креатин опт",
    "Протеин 900gr": "Протеин оптом",
    "Протеин Леврон Черный 2кг": "Протеин оптом",
    "Дада": "Сок Дада пластик бут 0.5",
    "Витамин D": "Витамин D3 180 капсул",
}
# Услуги UCS, которые не попадают в отчёт (абонементы, посещения и т.п.). * — любые символы.
DEFAULT_IGNORED = ("Акция * посещени*", "Безлимит*", "Молоко", "ПРОТЕИН В ПАЧКАХ ШТУЧНО", "BCA 300грамм большой")
DATE_FORMAT = r"dd\.mm\.yyyy"  # как в эталонном отчёте


class ReportFormatError(Exception):
    pass


def normalize(name: str) -> str:
    """Для сравнения названий: без учёта регистра, ё/е, запятой/точки и пробелов («30 гр» = «30гр»)."""
    s = str(name).lower().replace("ё", "е").replace(",", ".")
    return re.sub(r"\s+", "", s)


def is_ignored(name: str, patterns) -> bool:
    key = normalize(name)
    return any(fnmatchcase(key, normalize(p)) for p in patterns)


def report_filename(date: dt.date) -> str:
    return f"{REPORT_PREFIX}{date:%Y-%m-%d}.xlsx"


def report_date(name: str) -> dt.date | None:
    """Дата из имени дневного отчёта (оба формата) или None.
    NFC: «ё» в имени файла может быть записана двумя символами (е + ◌̈)."""
    m = REPORT_NAME_RE.match(unicodedata.normalize("NFC", name))
    if not m:
        return None
    fmt = "%Y-%m-%d" if "-" in m.group(1) else "%d.%m.%Y"
    return dt.datetime.strptime(m.group(1), fmt).date()


def find_reports(reports_dir: Path) -> dict[dt.date, Path]:
    """Все дневные отчёты в папке и её подпапках. Если за дату их два — берётся тот, что в папке месяца."""
    found: dict[dt.date, Path] = {}
    if not reports_dir.is_dir():
        return found
    for p in reports_dir.rglob("*.xlsx"):
        d = report_date(p.name)
        if d is None or not p.is_file():
            continue
        other = found.get(d)
        if other is None or (len(p.parts), "-" in p.name) > (len(other.parts), "-" in other.name):
            found[d] = p
    return found


def find_previous_report(reports_dir: Path, before: dt.date) -> Path | None:
    """Самый свежий дневной отчёт с датой раньше `before`."""
    reports = find_reports(reports_dir)
    earlier = [d for d in reports if d < before]
    return reports[max(earlier)] if earlier else None


@dataclass
class Layout:
    header_row: int
    cols: dict[str, int]          # название колонки -> номер (1-based)
    item_rows: dict[str, int]     # нормализованное название товара -> строка
    item_names: dict[str, str]    # нормализованное -> как написано в отчёте
    staff_rows: list[int]
    staff_keys: dict[str, int]    # нормализованное название строки персонала -> строка
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
    lay = Layout(header_row, cols, {}, {}, [], {}, None)
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
                lay.staff_keys[key] = r
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


def _match_sales(sales: dict[str, float], lay: Layout, aliases: dict[str, str], ignored) \
        -> tuple[dict[str, float], dict[str, str], list[str]]:
    """Продажи UCS по строкам отчёта: (продано по ключу, новые товары ключ -> название, пропущенные услуги)."""
    sold: dict[str, float] = {}
    new_names: dict[str, str] = {}
    alias = {normalize(k): v for k, v in aliases.items()}
    skipped: list[str] = []
    for name, qty in sales.items():
        if is_ignored(name, ignored):
            skipped.append(name)
            continue
        name = alias.get(normalize(name), name)
        key = normalize(name)
        sold[key] = sold.get(key, 0) + qty
        if key not in lay.item_rows and key not in lay.staff_keys and key not in new_names:
            new_names[key] = name
    return sold, new_names, skipped


def _fill_staff_sold(ws: Worksheet, lay: Layout, sold: dict[str, float], result: "BuildResult") -> None:
    """Продажи UCS, сопоставленные строкам персонала (например, «Утеря карт»), — в «Продано» этой строки."""
    for key, r in lay.staff_keys.items():
        s = sold.get(key)
        if s:
            ws.cell(r, lay.cols[COL_SOLD]).value = _clean(s)
            result.filled.append((str(ws.cell(r, NAME_COL).value).strip(), s))


def _evening_formula(c: dict[str, int], r: int) -> str:
    L = lambda col: get_column_letter(c[col])  # noqa: E731
    return f"={L(COL_MORNING)}{r}+{L(COL_INCOME)}{r}-{L(COL_SOLD)}{r}-{L(COL_WRITEOFF)}{r}"


@dataclass
class BuildResult:
    output: Path
    source: Path
    from_template: bool
    filled: list[tuple[str, float]] = field(default_factory=list)
    added: list[str] = field(default_factory=list)          # новые строки (названия из UCS как есть)
    negative: list[tuple[str, float]] = field(default_factory=list)
    ignored: list[str] = field(default_factory=list)        # услуги UCS, не попавшие в отчёт
    unsplit: list[str] = field(default_factory=list)        # услуги UCS, не разнесённые по цене
    cash_opening: float | None = None   # остаток наличных на начало дня; None — вносят вручную
    rebuilt: bool = False               # пересоздан по новой выгрузке, вписанное вручную сохранено
    opening_changed: list[str] = field(default_factory=list)  # при пересоздании: что обновилось из прошлого отчёта


def build_report(
    *,
    date: dt.date,
    sales: dict[str, float],
    source: Path,
    from_template: bool,
    output: Path,
    section_rows: tuple[str, ...] = DEFAULT_SECTION_ROWS,
    staff_rows: tuple[str, ...] = DEFAULT_STAFF_ROWS,
    aliases: dict[str, str] = DEFAULT_ALIASES,
    ignored: tuple[str, ...] = DEFAULT_IGNORED,
) -> BuildResult:
    """
    source — вчерашний отчёт (ОстУтро = вчерашний ОстВечер)
             или шаблон (ОстУтро берётся как есть, его вносят вручную при первом запуске).
    Названия из UCS сначала переводятся через aliases, затем продажа идёт в строку
    с тем же названием, а если такой строки нет — добавляется новая.
    """
    wb = openpyxl.load_workbook(source)
    wb_values = openpyxl.load_workbook(source, data_only=True)  # значения формул, если файл сохранялся в Excel
    ws = get_snack_sheet(wb)
    ws_values = get_snack_sheet(wb_values)
    prev_cash, prev_cash_values = get_cash_sheet(wb), get_cash_sheet(wb_values)
    lay = detect_layout(ws, section_rows, staff_rows)
    c = lay.cols

    # 1. Утренние остатки (по вчерашнему файлу, до вставки новых строк)
    morning: dict[str, float] = {}
    for key, r in lay.item_rows.items():
        if from_template:
            morning[key] = _num(ws.cell(r, c[COL_MORNING]).value)
        else:
            morning[key] = max(evening_value(ws, ws_values, c, r), 0.0)  # минус вчера — сегодня с нуля

    # 2. Продажи: суммируем по нормализованному названию, новые товары добавляем строками
    sold, new_names, skipped = _match_sales(sales, lay, aliases, ignored)
    result = BuildResult(output=output, source=source, from_template=from_template, added=list(new_names.values()),
                         ignored=skipped)
    if new_names:
        _insert_items(ws, lay, list(new_names.values()))
        lay = detect_layout(ws, section_rows, staff_rows)
        c = lay.cols

    # 3. Заполнение
    date_cell = ws.cell(max(lay.header_row - 1, 1), NAME_COL)
    date_cell.value = date
    date_cell.number_format = DATE_FORMAT

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
        ws.cell(r, c[COL_EVENING]).value = _evening_formula(c, r)
        if m - (s or 0) < 0:
            result.negative.append((lay.item_names[key], m - (s or 0)))

    # 4. Персонал и карты — каждый день с чистого листа
    for r in lay.staff_rows:
        for col in range(NAME_COL + 1, ws.max_column + 1):
            cell = ws.cell(r, col)
            if not isinstance(cell, MergedCell):
                cell.value = None
    _fill_staff_sold(ws, lay, sold, result)

    # 5. Кассовый лист: остаток наличных на начало дня — из вчерашнего отчёта
    opening = None
    if not from_template and prev_cash is not None:
        opening = closing_cash(prev_cash, prev_cash_values)
    cash_ws, _ = ensure_sheets(wb)
    fill_cash(cash_ws, date, opening)
    result.cash_opening = opening

    _reset_negative_highlight(ws, lay)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return result


def rebuild_report(
    *,
    date: dt.date,
    sales: dict[str, float],
    existing: Path,
    output: Path | None = None,
    section_rows: tuple[str, ...] = DEFAULT_SECTION_ROWS,
    staff_rows: tuple[str, ...] = DEFAULT_STAFF_ROWS,
    aliases: dict[str, str] = DEFAULT_ALIASES,
    ignored: tuple[str, ...] = DEFAULT_IGNORED,
    prev: Path | None = None,
) -> BuildResult:
    """Пересоздаёт уже готовый отчёт по новой выгрузке UCS.

    Меняется «Продано» (и добавляются строки новых товаров). Если есть предыдущий отчёт `prev`,
    ОстУтро и остаток наличных на начало дня заново берутся из него, даже если их правили вручную.
    Остальное вписанное вручную остаётся: Приход, Долг, Имя, Списание, Причина, строки персонала
    и весь кассовый лист. Без `prev` ОстУтро и наличные не трогаются.
    """
    output = output or existing
    wb = openpyxl.load_workbook(existing)
    opening_changed: list[str] = []
    if prev is not None:
        opening_changed, _ = _apply_opening(wb, prev, section_rows, staff_rows)
    ws = get_snack_sheet(wb)
    lay = detect_layout(ws, section_rows, staff_rows)
    sold, new_names, skipped = _match_sales(sales, lay, aliases, ignored)
    result = BuildResult(output=output, source=prev or existing, from_template=False,
                         added=list(new_names.values()), ignored=skipped, rebuilt=True,
                         opening_changed=opening_changed)
    if new_names:
        _insert_items(ws, lay, list(new_names.values()))
        lay = detect_layout(ws, section_rows, staff_rows)
    c = lay.cols

    g = lambda r, col: _num(ws.cell(r, c[col]).value) if col in c else 0.0  # noqa: E731
    for key, r in lay.item_rows.items():
        if key in new_names:
            ws.cell(r, c[COL_MORNING]).value = 0
        s = sold.get(key)
        ws.cell(r, c[COL_SOLD]).value = _clean(s) if s else None
        if s:
            result.filled.append((lay.item_names[key], s))
        ws.cell(r, c[COL_EVENING]).value = _evening_formula(c, r)
        evening = g(r, COL_MORNING) + g(r, COL_INCOME) - (s or 0) - g(r, COL_WRITEOFF)
        if evening < 0:
            result.negative.append((lay.item_names[key], evening))
    _fill_staff_sold(ws, lay, sold, result)

    cash_ws, _ = ensure_sheets(wb)
    opening = cash_ws[OPENING_CELL].value
    result.cash_opening = float(opening) if isinstance(opening, (int, float)) and not isinstance(opening, bool) \
        else None
    _reset_negative_highlight(ws, lay)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return result


def _apply_opening(
    wb,
    prev: Path,
    section_rows: tuple[str, ...],
    staff_rows: tuple[str, ...],
) -> tuple[list[str], bool]:
    """Ставит в книгу ОстУтро (= вчерашний ОстВечер) и остаток наличных на начало дня из `prev`.

    Товары, появившиеся в предыдущем отчёте, добавляются строками. Остальное не трогается.
    Возвращает (что поменялось, добавлены ли строки).
    """
    wb_prev = openpyxl.load_workbook(prev)
    wb_prev_values = openpyxl.load_workbook(prev, data_only=True)
    ws_p, ws_pv = get_snack_sheet(wb_prev), get_snack_sheet(wb_prev_values)
    lay_p = detect_layout(ws_p, section_rows, staff_rows)
    evening = {k: max(evening_value(ws_p, ws_pv, lay_p.cols, r), 0.0) for k, r in lay_p.item_rows.items()}

    ws = get_snack_sheet(wb)
    lay = detect_layout(ws, section_rows, staff_rows)
    missing = [lay_p.item_names[k] for k, _ in sorted(lay_p.item_rows.items(), key=lambda kv: kv[1])
               if k not in lay.item_rows]
    if missing:
        _insert_items(ws, lay, missing)
        lay = detect_layout(ws, section_rows, staff_rows)
    c = lay.cols
    new_keys = {normalize(n) for n in missing}

    changed: list[str] = []
    for key, r in lay.item_rows.items():
        if key not in evening:
            continue  # товар есть только в этом отчёте — не трогаем
        cell = ws.cell(r, c[COL_MORNING])
        if key in new_keys:
            for col in CLEARED_COLS:
                if col in c:
                    ws.cell(r, c[col]).value = None
            ws.cell(r, c[COL_EVENING]).value = _evening_formula(c, r)
        if key in new_keys or _num(cell.value) != evening[key]:
            cell.value = _clean(evening[key])
            changed.append(lay.item_names[key])

    prev_cash, cash = get_cash_sheet(wb_prev), get_cash_sheet(wb)
    if prev_cash is not None and cash is not None:
        opening = closing_cash(prev_cash, get_cash_sheet(wb_prev_values))
        cur = cash[OPENING_CELL].value
        if not isinstance(cur, (int, float)) or isinstance(cur, bool) or float(cur) != opening:
            cash[OPENING_CELL].value = int(opening) if float(opening).is_integer() else round(opening, 2)
            changed.append("остаток наличных на начало дня")
    return changed, bool(missing)


def refresh_opening(
    report: Path,
    prev: Path,
    section_rows: tuple[str, ...] = DEFAULT_SECTION_ROWS,
    staff_rows: tuple[str, ...] = DEFAULT_STAFF_ROWS,
) -> list[str]:
    """Обновляет ОстУтро и остаток наличных на начало дня по предыдущему отчёту (после пересоздания).

    Файл сохраняется, только если что-то поменялось. Возвращает, что поменялось.
    """
    wb = openpyxl.load_workbook(report)
    changed, added = _apply_opening(wb, prev, section_rows, staff_rows)
    if changed:
        if added:
            ws = get_snack_sheet(wb)
            _reset_negative_highlight(ws, detect_layout(ws, section_rows, staff_rows))
        wb.save(report)
    return changed
