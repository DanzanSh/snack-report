import datetime as dt
import shutil
from pathlib import Path

import openpyxl
import pytest

from conftest import ROOT
import snack_report.main as main_mod
from snack_report.cash_report import get_snack_sheet
from snack_report.layout import report_path
from snack_report.main import run
from snack_report.report_builder import find_previous_report, normalize

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"
# В выгрузке: Козинак МИНИ 1, Вода 0.5 3, Вода 1,5 л 1, соки грацио стекло 1
# «Вода 1,5 л» по встроенному сопоставлению пишется в строку «Вода 1,5»


def _rows(path):
    ws = get_snack_sheet(openpyxl.load_workbook(path))
    return ws, {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}


@pytest.fixture
def template(tmp_path):
    """Шаблон с заполненным ОстУтро и именами персонала (как после ручного ввода)."""
    t = tmp_path / "template.xlsx"
    shutil.copy(ROOT / "config" / "template.xlsx", t)
    wb = openpyxl.load_workbook(t)
    ws = get_snack_sheet(wb)
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
    assert normalize("Протеин 30 гр") == normalize("протеин 30гр")


def test_first_run_and_next_day(tmp_path, template, monkeypatch):
    out = tmp_path / "reports"
    orig = main_mod.read_ucs

    def with_new_item(p, *a, **kw):
        rep = orig(p, *a, **kw)
        rep.sales["Новинка"] = 1
        return rep

    monkeypatch.setattr(main_mod, "read_ucs", with_new_item)
    assert run([str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]) == 0
    monkeypatch.setattr(main_mod, "read_ucs", orig)
    day1 = report_path(out, dt.date(2026, 10, 28))
    ws, row = _rows(day1)
    assert ws["A1"].value.date() == dt.date(2026, 10, 28)

    # «Вода 0.5» совпала с «Вода 0,5» по названию
    r = row["Вода 0,5"]
    assert ws.cell(r, 3).value == 253 and ws.cell(r, 4).value == 3
    assert ws.cell(r, 9).value == f"=C{r}+B{r}-D{r}-G{r}"
    assert ws.cell(row["Банан"], 4).value is None
    assert ws["A1"].number_format == r"dd\.mm\.yyyy"

    # «Вода 1,5 л» из UCS записана в строку эталона «Вода 1,5», новой строки нет
    assert ws.cell(row["Вода 1,5"], 4).value == 1
    assert "Вода 1,5 л" not in row

    # «Козинак МИНИ» и «соки грацио стекло» — через сопоставление в строки отчёта
    assert ws.cell(row["Козинаки"], 4).value == 1
    assert ws.cell(row["Сок Грацио стекло 0.2"], 4).value == 1
    assert "Козинак МИНИ" not in row and "соки грацио стекло" not in row

    # Незнакомое название UCS добавлено как есть, перед «Прочее»
    n = "Новинка"
    assert row[n] == row["Прочее"] - 1 == row["Зефир 1000"] + 1
    assert ws.cell(row[n], 4).value == 1
    assert ws.cell(row[n], 9).value == f"=C{row[n]}+B{row[n]}-D{row[n]}-G{row[n]}"
    # формулы ниже вставки указывают на свои строки
    rp = row["Рис курица"]
    assert ws.cell(rp, 9).value == f"=C{rp}+B{rp}-D{rp}-G{rp}"

    # Имена персонала очищены
    assert ws.cell(row["Администратор"], 8).value is None
    assert ws.cell(row["Администратор"], 9).value is None

    # Администратор вносит приход и списание за 28.10 и имя смены
    ws.cell(r, 2).value = 10
    ws.cell(r, 7).value = 2
    ws.cell(row["Козинаки"], 2).value = 5
    ws.cell(row["Охрана"], 8).value = "Жанат"
    ws.parent.save(day1)

    # Следующий день: та же выгрузка, но за 29.10
    orig = main_mod.read_ucs

    def read_next_day(p, *_, **__):
        rep = orig(p)
        rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = read_next_day
    try:
        assert run([str(FIXTURE), "--console", "--out", str(out)]) == 0
    finally:
        main_mod.read_ucs = orig

    ws2, row2 = _rows(report_path(out, dt.date(2026, 10, 29)))
    r2 = row2["Вода 0,5"]
    assert ws2.cell(r2, 3).value == 253 + 10 - 3 - 2  # ОстУтро = вчерашний ОстВечер
    assert ws2.cell(r2, 2).value is None and ws2.cell(r2, 7).value is None
    assert ws2.cell(row2["Козинаки"], 3).value == 5 - 1
    assert len(row2) == len(row)  # повторно строки не добавляются
    assert ws2.cell(row2["Охрана"], 8).value is None


def test_dragging_export_again_rebuilds_report(tmp_path, template):
    out = tmp_path / "reports"
    args = [str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]
    assert run(args) == 0
    day = report_path(out, dt.date(2026, 10, 28))
    wb = openpyxl.load_workbook(day)
    ws = get_snack_sheet(wb)
    row = {ws.cell(r, 1).value: r for r in range(3, ws.max_row + 1)}
    ws.cell(row["Вода 0,5"], 2).value = 10   # Приход вписан вручную
    wb.save(day)

    assert run(args) == 0                    # без вопроса: вписанное вручную не теряется
    ws, row = _rows(day)
    assert ws.cell(row["Вода 0,5"], 2).value == 10
    assert ws.cell(row["Вода 0,5"], 4).value == 3


def test_find_previous_report(tmp_path):
    for n in ["Отчёт_27.10.2026.xlsx", "Отчёт_25.10.2026.xlsx", "Отчёт_29.10.2026.xlsx", "мусор.xlsx"]:
        (tmp_path / n).touch()
    assert find_previous_report(tmp_path, dt.date(2026, 10, 29)).name == "Отчёт_27.10.2026.xlsx"
    assert find_previous_report(tmp_path, dt.date(2026, 10, 1)) is None
    # «ё» из двух символов (так её иногда сохраняет файловый менеджер)
    (tmp_path / "Отче\u0308т_2026-10-28.xlsx").touch()
    assert find_previous_report(tmp_path, dt.date(2026, 10, 29)).name == "Отче\u0308т_2026-10-28.xlsx"


def test_folder_mode_processes_new_exports_in_date_order(tmp_path):
    folder = tmp_path / "report"
    folder.mkdir()
    # две выгрузки: за 28.10 и (подменённая дата) за 29.10, плюс посторонний .xls
    shutil.copy(FIXTURE, folder / "ucs_a.xls")
    shutil.copy(FIXTURE, folder / "ucs_b.xls")
    (folder / "мусор.xls").write_bytes(b"not excel")

    orig = main_mod.read_ucs

    def fake(p, *_, **__):
        rep = orig(p)
        if Path(p).name == "ucs_b.xls":
            rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = fake
    try:
        assert run(["--console", "--out", str(folder)]) == 0
        assert (report_path(folder, dt.date(2026, 10, 28))).exists()
        assert (report_path(folder, dt.date(2026, 10, 29))).exists()
        ws, row = _rows(report_path(folder, dt.date(2026, 10, 29)))
        assert ws.cell(row["Вода 0,5"], 3).value == 0  # 28.10 закончился на -3, 29.10 начинается с 0

        # повторный запуск ничего не пересоздаёт
        mtime = (report_path(folder, dt.date(2026, 10, 29))).stat().st_mtime_ns
        assert run(["--console", "--out", str(folder)]) == 0
        assert (report_path(folder, dt.date(2026, 10, 29))).stat().st_mtime_ns == mtime
    finally:
        main_mod.read_ucs = orig


def test_folder_mode_creates_missing_folder(tmp_path):
    folder = tmp_path / "report"
    assert run(["--console", "--out", str(folder)]) == 2
    assert folder.is_dir()


def test_negative_evening_gives_zero_morning(tmp_path, template):
    out = tmp_path / "reports"
    # Банан: ОстУтро 1, продаж нет; Вода 0,5: 253 − 3
    assert run([str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]) == 0
    day1 = report_path(out, dt.date(2026, 10, 28))
    ws, row = _rows(day1)
    ws.cell(row["Банан"], 7).value = 5            # списание больше остатка → ОстВечер = -4
    ws.cell(row["Вода 0,5"], 7).value = 2         # обычный случай: 253 - 3 - 2 = 248
    ws.parent.save(day1)

    orig = main_mod.read_ucs

    def read_next_day(p, *_, **__):
        rep = orig(p)
        rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = read_next_day
    try:
        assert run([str(FIXTURE), "--console", "--out", str(out)]) == 0
    finally:
        main_mod.read_ucs = orig

    ws2, row2 = _rows(report_path(out, dt.date(2026, 10, 29)))
    assert ws2.cell(row2["Банан"], 3).value == 0
    assert ws2.cell(row2["Вода 0,5"], 3).value == 248


def test_new_item_row_with_only_name_gives_zero_morning(tmp_path, template):
    out = tmp_path / "reports"
    assert run([str(FIXTURE), "--console", "--out", str(out), "--prev", str(template)]) == 0
    day1 = report_path(out, dt.date(2026, 10, 28))
    ws, row = _rows(day1)
    # позицию добавили вручную: заполнено только наименование
    at = row["Прочее"]
    ws.insert_rows(at)
    ws.cell(at, 1).value = "Новинка"
    ws.parent.save(day1)

    orig = main_mod.read_ucs

    def read_next_day(p, *_, **__):
        rep = orig(p)
        rep.date = dt.date(2026, 10, 29)
        return rep

    main_mod.read_ucs = read_next_day
    try:
        assert run([str(FIXTURE), "--console", "--out", str(out)]) == 0
    finally:
        main_mod.read_ucs = orig

    ws2, row2 = _rows(report_path(out, dt.date(2026, 10, 29)))
    assert ws2.cell(row2["Новинка"], 3).value == 0


def test_aliases_sum_into_report_row(tmp_path, template):
    from snack_report.report_builder import build_report

    out = tmp_path / "out.xlsx"
    res = build_report(
        date=dt.date(2026, 11, 4),
        sales={"протеиновый коктейль (на молоке)": 3, "протеиновый коктейль на воде": 1, "Протеин порц": 2,
               "Протеин 30гр": 4, "Акция 3 посещения": 1, "Безлимит 1 месяц до 17-00": 2,
               "Разовые посещения зал": 5, "Козинак МИНИ": 1, "Новинка": 1,
               "Гейнер на воде": 1, "Гейнер на Молоке": 2, "Протеин 900 gr": 1, "Молоко": 3,
               "ПРОТЕИН В ПАЧКАХ ШТУЧНО": 1, "BCA 300грамм большой": 1},
        source=template, from_template=True, output=out,
    )
    ws, row = _rows(out)
    assert ws.cell(row["Протеин порц"], 4).value == 6
    assert ws.cell(row["Протеин в пачках штучно"], 4).value == 4
    # абонементы не попадают в отчёт, «Разовые посещения зал» — строка отчёта
    assert res.ignored == ["Акция 3 посещения", "Безлимит 1 месяц до 17-00", "Молоко",
                           "ПРОТЕИН В ПАЧКАХ ШТУЧНО", "BCA 300грамм большой"]
    assert "Акция 3 посещения" not in row and "Безлимит 1 месяц до 17-00" not in row
    assert ws.cell(row["Разовые посещения зал"], 4).value == 5
    assert ws.cell(row["Козинаки"], 4).value == 1
    assert ws.cell(row["Гейнер"], 4).value == 3
    assert ws.cell(row["Протеин оптом"], 4).value == 1  # пробелы не важны: «900 gr» = «900gr»
    assert res.added == ["Новинка"]  # незнакомое название — новая строка


def test_lost_card_goes_to_staff_row(tmp_path, template):
    from snack_report.report_builder import build_report, rebuild_report

    out = tmp_path / "out.xlsx"
    res = build_report(date=dt.date(2026, 11, 5), sales={"утеря карточки": 2},
                       source=template, from_template=True, output=out)
    ws, row = _rows(out)
    assert ws.cell(row["Утеря карт"], 4).value == 2
    assert res.added == []  # строка персонала, а не новый товар
    assert ("Утеря карт", 2) in res.filled

    res = rebuild_report(date=dt.date(2026, 11, 5), sales={"утеря карточки": 3}, existing=out)
    ws, row = _rows(out)
    assert ws.cell(row["Утеря карт"], 4).value == 3
    assert res.added == []


def test_price_split_from_settings(tmp_path):
    (tmp_path / "settings.ini").write_text(
        "[price_split]\nРазовое что-то = 1000: Строка А; 1 500: Строка Б\n", encoding="utf-8")
    s = main_mod.load_settings(tmp_path)
    assert s["price_split"]["Разовое что-то"] == {1000.0: "Строка А", 1500.0: "Строка Б"}
    assert "Разовые зал Бокс" in s["price_split"]  # встроенное сохраняется


def test_aliases_from_settings(tmp_path):
    (tmp_path / "settings.ini").write_text(
        "[aliases]\nСникерс мини = Батончик 700\nВода 1 л = Вода 0,5\n", encoding="utf-8")
    s = main_mod.load_settings(tmp_path)
    assert s["aliases"]["Сникерс мини"] == "Батончик 700"
    assert s["aliases"]["Вода 1 л"] == "Вода 0,5"  # перекрывает встроенное
    assert s["aliases"]["RedBull 0.25"] == "Ред Булл"  # встроенные сохраняются
