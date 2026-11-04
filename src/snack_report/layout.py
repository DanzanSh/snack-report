"""Раскладка папки report: новые выгрузки в корне, всё обработанное — по папкам месяцев.

report\\
  Сводка_04.11.xls              <- новая выгрузка
  2026-10 Октябрь\\
    Месяц_10.2026.xlsx
    Отчёт_2026-10-28.xlsx
    UCS\\Сводка_28.10.xls        <- обработанные выгрузки
"""
from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path

from .report_builder import report_date, report_filename

MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль",
          "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
UCS_DIR = "UCS"
MONTHLY_NAME_RE = re.compile(r"^Месяц_(\d{2})\.(\d{4})\.xlsx$")


def monthly_filename(year: int, month: int) -> str:
    return f"Месяц_{month:02d}.{year}.xlsx"


def month_dir(root: Path, year: int, month: int) -> Path:
    return root / f"{year}-{month:02d} {MONTHS[month - 1].capitalize()}"


def report_path(root: Path, date: dt.date) -> Path:
    return month_dir(root, date.year, date.month) / report_filename(date)


def monthly_path(root: Path, year: int, month: int) -> Path:
    return month_dir(root, year, month) / monthly_filename(year, month)


def ucs_archive_dir(root: Path, date: dt.date) -> Path:
    return month_dir(root, date.year, date.month) / UCS_DIR


def move_into(src: Path, dst_dir: Path) -> Path:
    """Переносит файл в папку; при совпадении имени добавляет « (2)», « (3)»…"""
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    n = 2
    while dst.exists():
        dst = dst_dir / f"{src.stem} ({n}){src.suffix}"
        n += 1
    shutil.move(str(src), str(dst))
    return dst


def organize(root: Path) -> list[str]:
    """Раскладывает по папкам месяцев отчёты, лежащие в корне (раскладка старых версий).

    Повторный запуск ничего не делает. Возвращает описания файлов, которые перенести не удалось.
    """
    problems = []
    for p in sorted(root.iterdir()):
        if not p.is_file():
            continue
        d = report_date(p.name)
        if d is not None:
            dst = report_path(root, d)
        elif m := MONTHLY_NAME_RE.match(p.name):
            dst = monthly_path(root, int(m.group(2)), int(m.group(1)))
        else:
            continue
        if dst.exists():
            problems.append(f"{p.name}: в {dst.parent.name} уже есть {dst.name}, файл оставлен на месте")
            continue
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            p.rename(dst)
        except OSError as e:
            problems.append(f"{p.name}: не удалось перенести ({e})")
    return problems
