"""Чтение выгрузки UCS Shelter «Сводка по доходу» (.xls)."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

import xlrd

PERIOD_RE = re.compile(r"с\s+(\d{2}\.\d{2}\.\d{4})")


class UcsFormatError(Exception):
    """Файл не похож на выгрузку «Сводка по доходу»."""


@dataclass
class UcsReport:
    date: dt.date
    sales: dict[str, float] = field(default_factory=dict)  # название услуги -> кол-во


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


def _parse_sale_row(sheet, r: int) -> tuple[str, float] | None:
    """Строка продажи: текст (название), а сразу за ним первое число (кол-во).
    Заголовки групп и «Итого:» — без числа после текста или со словом «Итого»."""
    name = None
    for c in range(1, sheet.ncols):
        v = sheet.cell_value(r, c)
        if name is None:
            if isinstance(v, str) and v.strip():
                name = v.strip()
                if name.lower().startswith("итого"):
                    return None
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            return name, float(v)
        elif isinstance(v, str) and v.strip():
            return None
    return None


def _find_date(sheet) -> dt.date:
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            m = PERIOD_RE.search(_text(sheet, r, c))
            if m:
                return dt.datetime.strptime(m.group(1), "%d.%m.%Y").date()
    raise UcsFormatError("Не найден период отчёта («с ДД.ММ.ГГГГ по ...»)")


def read_ucs(path: str | Path) -> UcsReport:
    try:
        book = xlrd.open_workbook(str(path))
    except xlrd.XLRDError as e:
        raise UcsFormatError(f"Не удалось открыть {path}: {e}") from e
    sheet = book.sheet_by_index(0)

    header_row = _find_header(sheet)
    report = UcsReport(date=_find_date(sheet))

    in_services = False
    seen_services = False
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
        if sale is None:
            continue  # заголовок группы или «Итого:»
        name, qty = sale
        report.sales[name] = report.sales.get(name, 0) + qty

    if not seen_services:
        raise UcsFormatError("В файле нет блока «Услуги»")
    return report
