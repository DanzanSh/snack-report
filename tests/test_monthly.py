import datetime as dt
import os
import shutil
from pathlib import Path

import openpyxl
import pytest

import snack_report.main as main_mod
from snack_report.cash_report import get_snack_sheet
from snack_report.main import run

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"
# В каждой выгрузке: Козинак МИНИ 1, Вода 0.5 3, Вода 1,5 л 1, соки грацио стекло 1

DATES = {  # 31.10 пропущен, 02.11 пропущен
    "a.xls": dt.date(2026, 10, 29),
    "b.xls": dt.date(2026, 10, 30),
    "c.xls": dt.date(2026, 11, 1),
    "d.xls": dt.date(2026, 11, 3),
}


@pytest.fixture
def folder(tmp_path, monkeypatch):
    f = tmp_path / "report"
    f.mkdir()
    orig = main_mod.read_ucs

    def fake(p):
        rep = orig(p)
        rep.date = DATES[Path(p).name]
        return rep

    monkeypatch.setattr(main_mod, "read_ucs", fake)
    return f


def _add(folder, *names):
    for n in names:
        shutil.copy(FIXTURE, folder / n)


def _month(path):
    ws = get_snack_sheet(openpyxl.load_workbook(path))
    hdr = next(r for r in range(1, 10) if ws.cell(r, 1).value == "Товар")
    rows = {ws.cell(r, 1).value: [ws.cell(r, c).value for c in range(2, 8)] for r in range(hdr + 1, ws.max_row + 1)}
    return ws, rows


def _set_morning(path, name, value):
    wb = openpyxl.load_workbook(path)
    ws = get_snack_sheet(wb)
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 1).value == name:
            ws.cell(r, 3).value = value
    wb.save(path)


def test_monthly_totals_and_month_split(folder):
    _add(folder, "a.xls")
    assert run(["--console", "--out", str(folder)]) == 0
    # в первом отчёте администратор вписывает остатки на утро
    _set_morning(folder / "Отчёт_29.10.2026.xlsx", "Вода 0,5", 100)
    _set_morning(folder / "Отчёт_29.10.2026.xlsx", "Козинак МИНИ", 10)

    _add(folder, "b.xls", "c.xls", "d.xls")
    assert run(["--console", "--out", str(folder)]) == 0

    ws, oct_ = _month(folder / "Месяц_10.2026.xlsx")
    # ОстНачало, Приход, Продано, Долг, Списание, ОстКонец
    assert oct_["Вода 0,5"] == [100, 0, 6, 0, 0, 94]
    assert oct_["Козинак МИНИ"] == [10, 0, 2, 0, 0, 8]
    assert oct_["Банан"] == [0, 0, 0, 0, 0, 0]
    assert oct_["ИТОГО"][2].startswith("=SUM(")
    assert "Нет отчётов" not in str(ws["A3"].value)  # выходные и праздники — не пропуск

    ws, nov = _month(folder / "Месяц_11.2026.xlsx")
    assert nov["Вода 0,5"] == [94, 0, 6, 0, 0, 88]  # начало ноября = конец октября


def test_monthly_rebuilt_only_when_daily_changes(folder):
    _add(folder, "a.xls", "b.xls")
    assert run(["--console", "--out", str(folder)]) == 0
    month = folder / "Месяц_10.2026.xlsx"
    old = month.stat().st_mtime_ns

    assert run(["--console", "--out", str(folder)]) == 0
    assert month.stat().st_mtime_ns == old  # ничего не менялось — не пересобираем

    # администратор вносит приход в дневной отчёт
    day = folder / "Отчёт_30.10.2026.xlsx"
    wb = openpyxl.load_workbook(day)
    ws = get_snack_sheet(wb)
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 1).value == "Вода 0,5":
            ws.cell(r, 2).value = 24
            ws.cell(r, 5).value = 1
            ws.cell(r, 7).value = 2
    wb.save(day)
    t = month.stat().st_mtime + 5
    os.utime(day, (t, t))

    assert run(["--console", "--out", str(folder)]) == 0
    _, rows = _month(month)
    assert rows["Вода 0,5"] == [0, 24, 6, 1, 2, 18]  # 29.10 ушёл в -3, 30.10 начался с 0; Долг вычитается
