import datetime as dt
from pathlib import Path

import pytest

from snack_report.ucs_reader import UcsFormatError, read_ucs

FIXTURE = Path(__file__).parent / "fixtures" / "test.xls"


def test_reads_services_only():
    rep = read_ucs(FIXTURE)
    assert rep.date == dt.date(2026, 10, 28)
    assert rep.sales == {
        "Козинак МИНИ": 1,
        "Вода 0.5": 3,
        "Вода 1,5 л": 1,
        "соки грацио стекло": 1,
    }  # платежи (QR-код, наличные) в продажи не попадают
    assert (rep.cash, rep.noncash) == (20000, 174400)


def test_reads_payments_of_visits_export():
    rep = read_ucs(Path(__file__).parent / "fixtures" / "visits.xls")
    assert (rep.cash, rep.noncash) == (72000, 202350)


def test_not_an_ucs_file(tmp_path):
    bad = tmp_path / "bad.xls"
    bad.write_bytes(b"not excel")
    with pytest.raises(UcsFormatError):
        read_ucs(bad)


def test_skips_visits_group():
    rep = read_ucs(Path(__file__).parent / "fixtures" / "visits.xls")
    assert rep.sales == {
        "Батончик 1500": 2,
        "RedBull 0.25": 1,
        "Вода 0.5": 5,
        "Вода 1 л": 6,
        "Вода 1,5 л": 1,
        "Креатин порция": 1,
    }
    assert rep.ignored == [
        "12 посещений",
        "Абонемент MIX 1мес",
        "Безлимит 1 месяц",
        "Безлимит 1 месяц до 17-00",
        "Бойцовская зона группа 12пос",
    ]


def test_ignored_groups_configurable():
    rep = read_ucs(Path(__file__).parent / "fixtures" / "visits.xls", ignored_groups=())
    assert rep.sales["Абонемент MIX 1мес"] == 1
    assert rep.ignored == []


def test_split_by_price():
    from snack_report.ucs_reader import split_by_price

    prices = {4000.0: "Разовые посещения зал", 5000.0: "Разовое Бокс"}
    assert split_by_price(3, 12000, prices) == {"Разовые посещения зал": 3}
    assert split_by_price(2, 10000, prices) == {"Разовое Бокс": 2}
    assert split_by_price(3, 13000, prices) == {"Разовые посещения зал": 2, "Разовое Бокс": 1}
    assert split_by_price(2, 7000, prices) is None  # не раскладывается на 4000/5000


def test_single_visits_split_and_lost_card_kept(monkeypatch):
    """«Разовые зал Бокс» и «утеря карточки» берутся даже из группы «Посещения»."""
    import snack_report.ucs_reader as ur

    rows = [
        ["", "", "Посещения", "", ""],
        ["", "", "", "Разовые зал Бокс", 3.0, 4000.0, 13000.0],
        ["", "", "", "утеря карточки", 1.0, 1000.0, 1000.0],
        ["", "", "", "Безлимит 1 месяц", 1.0, 20000.0, 20000.0],
        ["", "", "", "Разовые зал Бокс", 2.0, 4000.0, 7000.0],
    ]

    class Sheet:
        def __init__(self, data):
            w = max(map(len, data))
            self.data = [r + [""] * (w - len(r)) for r in data]
            self.nrows, self.ncols = len(self.data), w

        def cell_value(self, r, c):
            return self.data[r][c]

        def row_values(self, r):
            return self.data[r]

    data = [["с 05.11.2026 06:00:00 по 05.11.2026 23:00:00"],
            ["Тип", "Наименование услуги", "Кол-во"],
            ["Услуги"], *rows, ['Итого по группе "Услуги":']]

    class Book:
        def sheet_by_index(self, i):
            return Sheet(data)

    monkeypatch.setattr(ur.xlrd, "open_workbook", lambda p: Book())
    rep = ur.read_ucs("x.xls", keep=["утеря карточки"])
    assert rep.sales == {"Разовые посещения зал": 2, "Разовое Бокс": 1, "утеря карточки": 1}
    assert rep.ignored == ["Безлимит 1 месяц"]
    assert rep.unsplit == ["Разовые зал Бокс: 2 шт на 7000"]
