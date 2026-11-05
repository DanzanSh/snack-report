"""Очистка раз в квартал: дневные отчёты и выгрузки UCS за прошлые кварталы удаляются, месячные отчёты остаются.

Месяц, в котором лежит самый свежий дневной отчёт, не чистится: из этого отчёта берётся ОстУтро
следующего дня, а месячный отчёт иначе пересобрался бы из одного дня.
Если пользователь отказался, в этом квартале больше не спрашиваем (файл STATE_FILE в папке report).
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

from .layout import MONTHS, UCS_DIR
from .report_builder import find_reports

STATE_FILE = ".cleanup"
MONTH_DIR_RE = re.compile(r"^(\d{4})-(\d{2}) ")


def quarter_start(d: dt.date) -> dt.date:
    return dt.date(d.year, 3 * ((d.month - 1) // 3) + 1, 1)


@dataclass
class CleanupPlan:
    files: list[Path] = field(default_factory=list)
    months: list[tuple[int, int]] = field(default_factory=list)   # (год, месяц) по порядку

    def describe_months(self) -> str:
        return ", ".join(f"{MONTHS[m - 1]} {y}" for y, m in self.months)


def find_old(root: Path, today: dt.date) -> CleanupPlan:
    """Что удалить: дневные отчёты и файлы в папках UCS за месяцы раньше текущего квартала."""
    start = quarter_start(today)
    reports = find_reports(root)
    keep = None
    if reports:
        latest = max(reports)
        keep = (latest.year, latest.month)

    def old(y: int, m: int) -> bool:
        return dt.date(y, m, 1) < start and (y, m) != keep

    files: list[Path] = []
    months: set[tuple[int, int]] = set()
    for d, p in reports.items():
        if old(d.year, d.month):
            files.append(p)
            months.add((d.year, d.month))
    if root.is_dir():
        for sub in root.iterdir():
            m = MONTH_DIR_RE.match(sub.name)
            ucs = sub / UCS_DIR
            if not (m and sub.is_dir() and ucs.is_dir()):
                continue
            y, mo = int(m.group(1)), int(m.group(2))
            if old(y, mo):
                found = [p for p in ucs.rglob("*") if p.is_file()]
                files += found
                if found:
                    months.add((y, mo))
    return CleanupPlan(files=sorted(files), months=sorted(months))


def run_cleanup(root: Path, plan: CleanupPlan) -> tuple[list[Path], list[str]]:
    """Удаляет файлы плана и опустевшие папки UCS. Возвращает (удалённые, ошибки)."""
    deleted, errors = [], []
    for p in plan.files:
        try:
            p.unlink()
            deleted.append(p)
        except OSError as e:
            errors.append(f"{p.name}: не удалось удалить ({e})")
    for p in {f.parent for f in plan.files}:
        if p.name == UCS_DIR and p.is_dir() and not any(p.iterdir()):
            try:
                p.rmdir()
            except OSError:
                pass
    return deleted, errors


def declined(root: Path, today: dt.date) -> bool:
    """В этом квартале уже ответили «Нет»."""
    try:
        return (root / STATE_FILE).read_text(encoding="utf-8").strip() == quarter_start(today).isoformat()
    except OSError:
        return False


def remember_declined(root: Path, today: dt.date) -> None:
    try:
        (root / STATE_FILE).write_text(quarter_start(today).isoformat(), encoding="utf-8")
    except OSError:
        pass
