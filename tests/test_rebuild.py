"""Пересоздание отчёта по новой выгрузке: вписанное вручную остаётся, следующие дни пересчитываются."""
import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

import snack_report.main as main_mod
from snack_report.cash_report import get_cash_sheet, get_snack_sheet
from snack_report.layout import report_path
from snack_report.main import run
from snack_report.ucs_reader import UcsReport

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"
D1, D2, D3 = dt.date(2026, 10, 28), dt.date(2026, 10, 29), dt.date(2026, 10, 30)
WATER = "Вода 0,5"
INCOME, MORNING, SOLD, DEBT, WHO, WRITEOFF = 2, 3, 4, 5, 6, 7


@pytest.fixture
def folder(tmp_path, monkeypatch):
    """(папка report, put): put(имя, дата, продажи) кладёт выгрузку с этими продажами."""
    f = tmp_path / "report"
    f.mkdir()
    exports: dict[str, UcsReport] = {}
    monkeypatch.setattr(main_mod, "read_ucs", lambda p: exports[Path(p).name])

    def put(name, date, sales):
        exports[name] = UcsReport(date, dict(sales))
        shutil.copy(FIXTURE, f / name)

    return f, put


def _open(path):
    wb = openpyxl.load_workbook(path)
    ws = get_snack_sheet(wb)
    return wb, ws, {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}


def _num(v):
    return v or 0


def test_rebuild_keeps_manual_data_and_updates_next_days(folder):
    folder, put = folder
    put("d1.xls", D1, {"Вода 0.5": 3})
    put("d2.xls", D2, {"Вода 0.5": 1})
    put("d3.xls", D3, {})
    assert run(["--console", "--out", str(folder)]) == 0
    p1, p2, p3 = (report_path(folder, d) for d in (D1, D2, D3))

    # день 1: вписываем всё вручную
    wb, ws, row = _open(p1)
    w = row[WATER]
    m1 = _num(ws.cell(w, MORNING).value)
    ws.cell(w, INCOME).value = 10
    ws.cell(w, DEBT).value = 2
    ws.cell(w, WHO).value = "Иван"
    ws.cell(w, WRITEOFF).value = 1
    ws.cell(row["Администратор"], 8).value = "Светлана"
    cash = get_cash_sheet(wb)
    c5_1 = _num(cash["C5"].value)
    cash["C17"].value = 5000
    cash["B8"].value, cash["C8"].value = "Вода для кулера", 200
    cash["D3"].value = "Светлана"
    wb.save(p1)

    # день 2: свой приход и выручка
    wb, ws, row = _open(p2)
    ws.cell(row["Банан"], INCOME).value = 4
    get_cash_sheet(wb)["C17"].value = 100
    wb.save(p2)

    # новая выгрузка за день 1: продаж больше и новый товар
    put("d1.xls", D1, {"Вода 0.5": 5, "Новый батончик": 2})
    assert run(["--console", "--out", str(folder)]) == 0
    assert not (folder / "d1.xls").exists()

    wb, ws, row = _open(p1)
    w = row[WATER]
    assert ws.cell(w, SOLD).value == 5
    assert [ws.cell(w, c).value for c in (INCOME, DEBT, WHO, WRITEOFF)] == [10, 2, "Иван", 1]
    assert _num(ws.cell(w, MORNING).value) == m1
    assert ws.cell(row["Администратор"], 8).value == "Светлана"
    assert ws.cell(row["Новый батончик"], MORNING).value == 0
    assert ws.cell(row["Новый батончик"], SOLD).value == 2
    cash = get_cash_sheet(wb)
    assert (cash["C17"].value, cash["B8"].value, cash["C8"].value, cash["D3"].value) == \
        (5000, "Вода для кулера", 200, "Светлана")

    # день 2: ОстУтро и наличные на начало — из нового дня 1, вписанное вручную на месте
    wb, ws, row = _open(p2)
    w = row[WATER]
    e1 = max(m1 + 10 - 5 - 1, 0)
    assert _num(ws.cell(w, MORNING).value) == e1
    assert ws.cell(w, SOLD).value == 1
    assert ws.cell(row["Банан"], INCOME).value == 4
    assert ws.cell(row["Новый батончик"], MORNING).value == 0   # вчера ушёл в минус
    cash = get_cash_sheet(wb)
    c5_2 = c5_1 + 5000 - 1000 - 200
    assert cash["C5"].value == c5_2 and cash["C17"].value == 100

    # день 3 — по цепочке
    wb, ws, row = _open(p3)
    assert _num(ws.cell(row[WATER], MORNING).value) == max(e1 - 1, 0)
    assert get_cash_sheet(wb)["C5"].value == c5_2 + 100 - 1000


def test_rebuild_clears_old_sales(folder):
    folder, put = folder
    put("d1.xls", D1, {"Вода 0.5": 3, "Банан": 1})
    assert run(["--console", "--out", str(folder)]) == 0
    put("d1.xls", D1, {"Банан": 2})
    assert run(["--console", "--out", str(folder)]) == 0
    _, ws, row = _open(report_path(folder, D1))
    assert ws.cell(row[WATER], SOLD).value is None
    assert ws.cell(row["Банан"], SOLD).value == 2


def test_report_open_in_excel_keeps_export(folder, monkeypatch):
    folder, put = folder
    put("d1.xls", D1, {"Вода 0.5": 3})
    assert run(["--console", "--out", str(folder)]) == 0
    put("d1.xls", D1, {"Вода 0.5": 4})

    def locked(*a, **k):
        raise PermissionError("locked")

    monkeypatch.setattr(main_mod, "rebuild_report", locked)
    assert run(["--console", "--out", str(folder)]) == 0
    assert (folder / "d1.xls").exists()   # остаётся в корне, обработается при следующем запуске
