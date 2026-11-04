import datetime as dt
from pathlib import Path

import openpyxl
import pytest

import snack_report.main as main_mod
from snack_report.cash_report import CASH_SHEET, SNACK_SHEET, get_snack_sheet
from snack_report.main import run

from conftest import ROOT

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"


@pytest.fixture
def day(monkeypatch):
    """Запуск по выгрузке с подменой даты."""
    orig = main_mod.read_ucs

    def go(args, date):
        def fake(p):
            rep = orig(p)
            rep.date = date
            return rep

        monkeypatch.setattr(main_mod, "read_ucs", fake)
        return run(["--console", *args])

    return go


def _first_day(day, out):
    template = ROOT / "config" / "template.xlsx"   # встроенный шаблон: первый отчёт
    assert day([str(FIXTURE), "--out", str(out), "--prev", str(template)], dt.date(2026, 10, 28)) == 0
    return out / "Отчёт_28.10.2026.xlsx"


def test_two_sheets_and_fixed_fields(tmp_path, day):
    path = _first_day(day, tmp_path / "r")
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == [CASH_SHEET, SNACK_SHEET]
    assert all(len(n) <= 31 for n in wb.sheetnames)
    cash = wb[CASH_SHEET]
    assert cash["B3"].value.date() == dt.date(2026, 10, 28)
    assert cash["B7"].value == "Дорожные" and cash["C7"].value == 1000
    assert cash["C23"].value == "=C5+C17+C20+C21-C15"
    assert cash["C5"].value is None            # первый отчёт: остаток вносят вручную
    assert wb[SNACK_SHEET]["A1"].value.date() == dt.date(2026, 10, 28)
    assert sum(ws.sheet_view.tabSelected for ws in wb) == 1


def test_next_day_carries_cash_and_resets_the_rest(tmp_path, day):
    out = tmp_path / "r"
    day1 = _first_day(day, out)
    wb = openpyxl.load_workbook(day1)
    c = wb[CASH_SHEET]
    c["C5"] = 10000
    c["C17"], c["C18"] = 5000, 7000
    c["B8"], c["C8"], c["D8"] = "молоко", 200, "для кофе"
    c["C20"] = 300
    c["C7"] = 1500                      # «Дорожные» правили вручную — завтра снова 1000
    c["D3"] = "Администратор Калина"
    c["C25"], c["C26"] = "Гулия", "Кемелбай"
    wb.save(day1)

    assert day([str(FIXTURE), "--out", str(out)], dt.date(2026, 10, 29)) == 0
    c2 = openpyxl.load_workbook(out / "Отчёт_29.10.2026.xlsx")[CASH_SHEET]
    assert c2["C5"].value == 10000 + 5000 + 300 - (1500 + 200)   # конец дня 28.10
    assert c2["B3"].value.date() == dt.date(2026, 10, 29)
    assert c2["C7"].value == 1000
    assert c2["B8"].value is None and c2["C8"].value is None and c2["D8"].value is None
    assert c2["C17"].value is None and c2["C18"].value is None and c2["C20"].value is None
    assert c2["D3"].value.strip() == "Администратор"
    assert c2["C25"].value is None and c2["C26"].value is None


def test_old_single_sheet_report_gets_cash_sheet(tmp_path, day):
    out = tmp_path / "r"
    day1 = _first_day(day, out)
    wb = openpyxl.load_workbook(day1)
    del wb[CASH_SHEET]
    wb[SNACK_SHEET].title = "Отчёт"       # формат до появления кассового листа
    wb.save(day1)

    assert day([str(FIXTURE), "--out", str(out)], dt.date(2026, 10, 29)) == 0
    wb2 = openpyxl.load_workbook(out / "Отчёт_29.10.2026.xlsx")
    assert wb2.sheetnames == [CASH_SHEET, SNACK_SHEET]
    assert wb2[CASH_SHEET]["C7"].value == 1000 and wb2[CASH_SHEET]["C5"].value is None
    assert get_snack_sheet(wb2).title == SNACK_SHEET
