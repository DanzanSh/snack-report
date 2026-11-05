import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

import snack_report.main as main_mod
from snack_report.cash_report import get_snack_sheet
from snack_report.layout import month_dir, monthly_path, move_into, organize, report_path
from snack_report.main import run
from snack_report.report_builder import find_previous_report, find_reports, report_date, report_filename

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"
DATES = {"a.xls": dt.date(2026, 10, 31), "b.xls": dt.date(2026, 11, 2)}


@pytest.fixture
def folder(tmp_path, monkeypatch):
    f = tmp_path / "report"
    f.mkdir()
    orig = main_mod.read_ucs

    def fake(p, *_, **__):
        rep = orig(p)
        rep.date = DATES.get(Path(p).name, rep.date)
        return rep

    monkeypatch.setattr(main_mod, "read_ucs", fake)
    return f


def test_names_and_paths(tmp_path):
    d = dt.date(2026, 10, 28)
    assert report_filename(d) == "Отчёт_2026-10-28.xlsx"
    assert report_date("Отчёт_2026-10-28.xlsx") == d
    assert report_date("Отчёт_28.10.2026.xlsx") == d   # старый формат
    assert report_date("~$Отчёт_2026-10-28.xlsx") is None
    assert report_path(tmp_path, d) == tmp_path / "2026-10 Октябрь" / "Отчёт_2026-10-28.xlsx"
    assert monthly_path(tmp_path, 2026, 10) == tmp_path / "2026-10 Октябрь" / "Месяц_10.2026.xlsx"


def test_previous_report_across_months_and_formats(tmp_path):
    report_path(tmp_path, dt.date(2026, 10, 31)).parent.mkdir()
    report_path(tmp_path, dt.date(2026, 10, 31)).touch()
    (tmp_path / "Отчёт_30.10.2026.xlsx").touch()       # старое имя в корне
    assert find_previous_report(tmp_path, dt.date(2026, 11, 1)) == report_path(tmp_path, dt.date(2026, 10, 31))
    assert find_previous_report(tmp_path, dt.date(2026, 10, 31)).name == "Отчёт_30.10.2026.xlsx"


def test_folder_run_archives_exports(folder):
    shutil.copy(FIXTURE, folder / "a.xls")
    shutil.copy(FIXTURE, folder / "b.xls")
    (folder / "мусор.xls").write_bytes(b"not excel")
    assert run(["--console", "--out", str(folder)]) == 0

    oct_, nov = month_dir(folder, 2026, 10), month_dir(folder, 2026, 11)
    assert (oct_ / "UCS" / "a.xls").exists() and (nov / "UCS" / "b.xls").exists()
    assert sorted(p.name for p in folder.iterdir() if p.is_file()) == ["мусор.xls"]  # нечитаемый остаётся
    assert (oct_ / "Месяц_10.2026.xlsx").exists() and (nov / "Месяц_11.2026.xlsx").exists()

    # 02.11 берёт остатки из 31.10 — из папки прошлого месяца
    ws = get_snack_sheet(openpyxl.load_workbook(report_path(folder, dt.date(2026, 11, 2))))
    row = {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}
    assert ws.cell(row["Вода 0,5"], 3).value == 0     # 31.10 закончился на -3

    # повторный запуск с пустым корнем — не ошибка
    assert run(["--console", "--out", str(folder)]) == 0


def test_export_for_existing_report_is_rebuilt_and_archived(folder):
    shutil.copy(FIXTURE, folder / "a.xls")
    assert run(["--console", "--out", str(folder)]) == 0
    day = report_path(folder, dt.date(2026, 10, 31))

    shutil.copy(FIXTURE, folder / "a.xls")             # ту же выгрузку положили ещё раз
    assert run(["--console", "--out", str(folder)]) == 0
    assert (day.parent / "UCS" / "a (2).xls").exists()
    assert not (folder / "a.xls").exists()
    ws = get_snack_sheet(openpyxl.load_workbook(day))
    row = {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}
    assert ws.cell(row["Вода 0,5"], 4).value == 3
    assert list(row).count("Козинак МИНИ") == 1        # новые строки не задвоились


def test_old_flat_folder_is_migrated(folder):
    shutil.copy(FIXTURE, folder / "a.xls")
    assert run(["--console", "--out", str(folder)]) == 0
    # имитируем раскладку старой версии: отчёт и месячный — плоско в корне, старое имя
    day = report_path(folder, dt.date(2026, 10, 31))
    day.rename(folder / "Отчёт_31.10.2026.xlsx")
    monthly_path(folder, 2026, 10).rename(folder / "Месяц_10.2026.xlsx")

    assert organize(folder) == []
    assert day.exists() and monthly_path(folder, 2026, 10).exists()
    assert not list(folder.glob("*.xlsx"))
    assert organize(folder) == []                       # повторно — ничего не делает


def test_migration_keeps_file_when_target_exists(tmp_path):
    target = report_path(tmp_path, dt.date(2026, 10, 31))
    target.parent.mkdir()
    target.write_text("new")
    (tmp_path / "Отчёт_31.10.2026.xlsx").write_text("old")
    problems = organize(tmp_path)
    assert len(problems) == 1 and (tmp_path / "Отчёт_31.10.2026.xlsx").exists()
    assert find_reports(tmp_path)[dt.date(2026, 10, 31)] == target   # приоритет у папки месяца


def test_move_into_renames_on_collision(tmp_path):
    (tmp_path / "x.xls").write_text("1")
    dst = tmp_path / "arch"
    assert move_into(tmp_path / "x.xls", dst) == dst / "x.xls"
    (tmp_path / "x.xls").write_text("2")
    assert move_into(tmp_path / "x.xls", dst) == dst / "x (2).xls"
