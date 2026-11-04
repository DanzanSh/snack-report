import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

from conftest import ROOT
import snack_report.main as main_mod
from snack_report.main import run
from snack_report.report_builder import find_previous_report, normalize

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"
# В выгрузке: Козинак МИНИ 1, Вода 0.5 3, Вода 1,5 л 1, соки грацио стекло 1


def _rows(path):
    ws = openpyxl.load_workbook(path).worksheets[0]
    return ws, {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}


@pytest.fixture
def template(tmp_path):
    """Шаблон с заполненным ОстУтро и именами персонала (как после ручного ввода)."""
    t = tmp_path / "template.xlsx"
    shutil.copy(ROOT / "config" / "template.xlsx", t)
    wb = openpyxl.load_workbook(t)
    ws = wb.worksheets[0]
    stock = {"Вода 0,5": 253, "Банан": 1}
    for r in range(3, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name in stock:
            ws.cell(r, 3).value = stock[name]
        if name == "Администратор":
            ws.cell(r, 8).value = "Светлана"
    wb.save(t)
    return t


def test_normalize():
    assert normalize("  Вода  1,5 ") == normalize("вода 1.5")
    assert normalize("Ёлка") == "елка"


def test_first_run_and_next_day(tmp_path, template):
    out = tmp_path / "reports"
    assert run([str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]) == 0
    day1 = out / "Отчёт_28.10.2026.xlsx"
    ws, row = _rows(day1)
    assert ws["A1"].value.date() == dt.date(2026, 10, 28)

    # «Вода 0.5» совпала с «Вода 0,5» по названию
    r = row["Вода 0,5"]
    assert ws.cell(r, 3).value == 253 and ws.cell(r, 4).value == 3
    assert ws.cell(r, 9).value == f"=C{r}+B{r}-D{r}-G{r}"
    assert ws.cell(row["Банан"], 4).value is None

    # Остальные названия UCS добавлены как есть, перед «Прочее», а «Козинаки» не тронуты
    new = ["Козинак МИНИ", "Вода 1,5 л", "соки грацио стекло"]
    assert [row[n] for n in new] == [row["Прочее"] - 3, row["Прочее"] - 2, row["Прочее"] - 1]
    assert row[new[0]] == row["Зефир 1000"] + 1
    for n in new:
        assert ws.cell(row[n], 4).value == 1
        assert ws.cell(row[n], 9).value == f"=C{row[n]}+B{row[n]}-D{row[n]}-G{row[n]}"
    assert ws.cell(row["Козинаки"], 4).value is None
    # формулы ниже вставки указывают на свои строки
    rp = row["Рис курица"]
    assert ws.cell(rp, 9).value == f"=C{rp}+B{rp}-D{rp}-G{rp}"

    # Имена персонала очищены
    assert ws.cell(row["Администратор"], 8).value is None
    assert ws.cell(row["Администратор"], 9).value is None

    # Администратор вносит приход и списание за 28.10 и имя смены
    ws.cell(r, 2).value = 10
    ws.cell(r, 7).value = 2
    ws.cell(row["Козинак МИНИ"], 2).value = 5
    ws.cell(row["Охрана"], 8).value = "Жанат"
    ws.parent.save(day1)

    # Следующий день: та же выгрузка, но за 29.10
    orig = main_mod.read_ucs

    def read_next_day(p):
        rep = orig(p)
        rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = read_next_day
    try:
        assert run([str(FIXTURE), "--console", "--out", str(out)]) == 0
    finally:
        main_mod.read_ucs = orig

    ws2, row2 = _rows(out / "Отчёт_29.10.2026.xlsx")
    r2 = row2["Вода 0,5"]
    assert ws2.cell(r2, 3).value == 253 + 10 - 3 - 2  # ОстУтро = вчерашний ОстВечер
    assert ws2.cell(r2, 2).value is None and ws2.cell(r2, 7).value is None
    assert ws2.cell(row2["Козинак МИНИ"], 3).value == 5 - 1
    assert len(row2) == len(row)  # повторно строки не добавляются
    assert ws2.cell(row2["Охрана"], 8).value is None


def test_existing_report_not_overwritten(tmp_path, template):
    out = tmp_path / "reports"
    args = [str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]
    assert run(args) == 0
    assert run(args) == 3
    assert run(args + ["--force"]) == 0


def test_find_previous_report(tmp_path):
    for n in ["Отчёт_27.10.2026.xlsx", "Отчёт_25.10.2026.xlsx", "Отчёт_29.10.2026.xlsx", "мусор.xlsx"]:
        (tmp_path / n).touch()
    assert find_previous_report(tmp_path, dt.date(2026, 10, 29)).name == "Отчёт_27.10.2026.xlsx"
    assert find_previous_report(tmp_path, dt.date(2026, 10, 1)) is None


def test_folder_mode_processes_new_exports_in_date_order(tmp_path):
    folder = tmp_path / "report"
    folder.mkdir()
    # две выгрузки: за 28.10 и (подменённая дата) за 29.10, плюс посторонний .xls
    shutil.copy(FIXTURE, folder / "ucs_a.xls")
    shutil.copy(FIXTURE, folder / "ucs_b.xls")
    (folder / "мусор.xls").write_bytes(b"not excel")

    orig = main_mod.read_ucs

    def fake(p):
        rep = orig(p)
        if Path(p).name == "ucs_b.xls":
            rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = fake
    try:
        assert run(["--console", "--out", str(folder)]) == 0
        assert (folder / "Отчёт_28.10.2026.xlsx").exists()
        assert (folder / "Отчёт_29.10.2026.xlsx").exists()
        ws, row = _rows(folder / "Отчёт_29.10.2026.xlsx")
        assert ws.cell(row["Вода 0,5"], 3).value == -3  # остаток перенесён из 28.10 (шаблон пустой)

        # повторный запуск ничего не пересоздаёт
        mtime = (folder / "Отчёт_29.10.2026.xlsx").stat().st_mtime_ns
        assert run(["--console", "--out", str(folder)]) == 0
        assert (folder / "Отчёт_29.10.2026.xlsx").stat().st_mtime_ns == mtime
    finally:
        main_mod.read_ucs = orig


def test_folder_mode_creates_missing_folder(tmp_path):
    folder = tmp_path / "report"
    assert run(["--console", "--out", str(folder)]) == 2
    assert folder.is_dir()
