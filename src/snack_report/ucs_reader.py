"""Чтение выгрузки UCS Shelter «Сводка по доходу» (.xls)."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

import xlrd

PERIOD_RE = re.compile(r"с\s+(\d{2}\.\d{2}\.\d{4})")
# Группы услуг UCS, которые целиком не попадают в отчёт (абонементы и посещения)
DEFAULT_IGNORED_GROUPS = ("Посещения",)
# Услуга UCS, которая по цене за штуку разносится по разным строкам отчёта: название -> {цена: строка отчёта}.
# Работает и внутри пропускаемых групп.
DEFAULT_PRICE_SPLIT = {
    "Разовые зал Бокс": {4000.0: "Разовые посещения зал", 5000.0: "Разовое Бокс"},
}


class UcsFormatError(Exception):
    """Файл не похож на выгрузку «Сводка по доходу»."""


@dataclass
class UcsReport:
    date: dt.date
    sales: dict[str, float] = field(default_factory=dict)  # название услуги -> кол-во
    ignored: list[str] = field(default_factory=list)  # услуги из пропущенных групп
    unsplit: list[str] = field(default_factory=list)  # не удалось разнести по цене (описание строки)


def _text(sheet, r: int, c: int) -> str:
    return str(sheet.cell_value(r, c)).strip()


def _find_header(sheet) -> int:
    """Строка заголовка таблицы. Колонки заголовка в выгрузке не совпадают
    с колонками данных (объединённые ячейки), поэтому берём только номер строки."""
    for r in range(sheet.nrows):
        row = [_text(sheet, r, c).lower() for c in range(sheet.ncols)]
        if "наименование услуги" in row and "кол-во" in row:
            return r
    raise UcsFormatError("Не найдена строка заголовка с колонками «Наименование услуги» и «Кол-во»")


def _parse_sale_row(sheet, r: int) -> tuple[str, float, float] | None:
    """Строка продажи: текст (название), за ним числа: кол-во, цена, сумма (сумма — последнее число).
    Заголовки групп и «Итого:» — без числа после текста или со словом «Итого»."""
    name = None
    numbers: list[float] = []
    for c in range(1, sheet.ncols):
        v = sheet.cell_value(r, c)
        if name is None:
            if isinstance(v, str) and v.strip():
                name = v.strip()
                if name.lower().startswith("итого"):
                    return None
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            numbers.append(float(v))
        elif isinstance(v, str) and v.strip():
            break
    if name is None or not numbers:
        return None
    return name, numbers[0], numbers[-1]


def split_by_price(qty: float, total: float, prices: dict[float, str]) -> dict[str, float] | None:
    """Разносит qty штук на сумму total по строкам отчёта с разной ценой.
    Одна цена на всё — всё в её строку; две цены — решаем a + b = qty, a*p1 + b*p2 = total.
    None — если сумма не раскладывается на целые штуки."""
    if qty <= 0:
        return None
    for price, row in prices.items():
        if abs(price * qty - total) < 0.01:
            return {row: qty}
    ordered = sorted(prices.items())
    for i, (p1, row1) in enumerate(ordered):
        for p2, row2 in ordered[i + 1:]:
            b = (total - p1 * qty) / (p2 - p1)
            if 0 < b < qty and abs(b - round(b)) < 1e-6:
                b = float(round(b))
                return {row1: qty - b, row2: b}
    return None


def _group_name(sheet, r: int) -> str | None:
    """Заголовок группы: строка с одним текстом без чисел (не «Итого:»)."""
    texts = [v.strip() for v in sheet.row_values(r)[1:] if isinstance(v, str) and v.strip()]
    if len(texts) == 1 and not texts[0].lower().startswith("итого"):
        return texts[0]
    return None


def _find_date(sheet) -> dt.date:
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            m = PERIOD_RE.search(_text(sheet, r, c))
            if m:
                return dt.datetime.strptime(m.group(1), "%d.%m.%Y").date()
    raise UcsFormatError("Не найден период отчёта («с ДД.ММ.ГГГГ по ...»)")


def _key(name: str) -> str:
    return "".join(str(name).lower().replace("ё", "е").split())


def read_ucs(path: str | Path, ignored_groups=DEFAULT_IGNORED_GROUPS, price_split=None,
             keep=()) -> UcsReport:
    """keep — названия услуг, которые берутся даже из пропускаемых групп (например, у которых есть сопоставление)."""
    try:
        book = xlrd.open_workbook(str(path))
    except xlrd.XLRDError as e:
        raise UcsFormatError(f"Не удалось открыть {path}: {e}") from e
    sheet = book.sheet_by_index(0)

    header_row = _find_header(sheet)
    report = UcsReport(date=_find_date(sheet))

    skip_keys = {g.strip().lower() for g in ignored_groups}
    splits = {_key(k): v for k, v in (DEFAULT_PRICE_SPLIT if price_split is None else price_split).items()}
    keep_keys = {_key(k) for k in keep}
    in_services = False
    seen_services = False
    group = ""
    for r in range(header_row + 1, sheet.nrows):
        first = _text(sheet, r, 0)
        if first == "Услуги":
            in_services = seen_services = True
            continue
        if first.startswith("Итого по группе"):
            if in_services:
                break
            continue
        if first:  # начался другой блок (например, «Платежи»)
            in_services = False
            continue
        if not in_services:
            continue

        sale = _parse_sale_row(sheet, r)
        if sale is None:  # заголовок группы или «Итого:»
            group = _group_name(sheet, r) or group
            continue
        name, qty, total = sale
        key = _key(name)
        if key in splits:
            parts = split_by_price(qty, total, splits[key])
            if parts is None:
                report.unsplit.append(f"{name}: {qty:g} шт на {total:g}")
                continue
            for row, n in parts.items():
                report.sales[row] = report.sales.get(row, 0) + n
            continue
        if group.lower() in skip_keys and key not in keep_keys:
            if name not in report.ignored:
                report.ignored.append(name)
            continue
        report.sales[name] = report.sales.get(name, 0) + qty

    if not seen_services:
        raise UcsFormatError("В файле нет блока «Услуги»")
    return report
