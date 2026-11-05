import datetime as dt

import pytest

import snack_report.main as main_mod
from snack_report.cleanup import declined, find_old, quarter_start, remember_declined, run_cleanup
from snack_report.layout import monthly_path, report_path, ucs_archive_dir


@pytest.mark.parametrize("day, start", [
    (dt.date(2026, 1, 1), dt.date(2026, 1, 1)),
    (dt.date(2026, 3, 31), dt.date(2026, 1, 1)),
    (dt.date(2026, 5, 15), dt.date(2026, 4, 1)),
    (dt.date(2026, 9, 30), dt.date(2026, 7, 1)),
    (dt.date(2026, 11, 5), dt.date(2026, 10, 1)),
])
def test_quarter_start(day, start):
    assert quarter_start(day) == start


def _touch(p):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.touch()
    return p


@pytest.fixture
def root(tmp_path):
    r = tmp_path / "report"
    for d in (dt.date(2026, 8, 3), dt.date(2026, 9, 30), dt.date(2026, 10, 1), dt.date(2026, 11, 4)):
        _touch(report_path(r, d))
        _touch(ucs_archive_dir(r, d) / f"{d:%d.%m}.xls")
        _touch(monthly_path(r, d.year, d.month))
    return r


def test_old_quarters_removed_monthly_kept(root):
    plan = find_old(root, dt.date(2026, 11, 5))
    assert plan.months == [(2026, 8), (2026, 9)]
    assert plan.describe_months() == "август 2026, сентябрь 2026"
    deleted, errors = run_cleanup(root, plan)
    assert errors == [] and len(deleted) == 4
    for d in (dt.date(2026, 8, 3), dt.date(2026, 9, 30)):
        assert not report_path(root, d).exists()
        assert not ucs_archive_dir(root, d).exists()
        assert monthly_path(root, d.year, d.month).exists()
    for d in (dt.date(2026, 10, 1), dt.date(2026, 11, 4)):
        assert report_path(root, d).exists() and ucs_archive_dir(root, d).exists()
    assert find_old(root, dt.date(2026, 11, 5)).files == []


def test_month_with_latest_report_is_kept(root):
    # новый квартал начался, а отчётов в нём ещё нет: ноябрь — последний, его не трогаем
    plan = find_old(root, dt.date(2027, 1, 2))
    assert plan.months == [(2026, 8), (2026, 9), (2026, 10)]
    assert report_path(root, dt.date(2026, 11, 4)) not in plan.files


def test_decline_remembered_until_next_quarter(root):
    today = dt.date(2026, 11, 5)
    assert not declined(root, today)
    remember_declined(root, today)
    assert declined(root, dt.date(2026, 12, 31))
    assert not declined(root, dt.date(2027, 1, 1))


class FakeUi:
    gui = True

    def __init__(self, answer):
        self.answer, self.asked = answer, []

    def confirm(self, text):
        self.asked.append(text)
        return self.answer


def test_main_asks_once_per_quarter(root, monkeypatch):
    monkeypatch.setattr(main_mod, "_today", lambda: dt.date(2026, 11, 5))
    ui = FakeUi(False)
    assert main_mod._cleanup(root, ui) == []
    assert len(ui.asked) == 1 and "август 2026, сентябрь 2026" in ui.asked[0]
    assert main_mod._cleanup(root, ui) == [] and len(ui.asked) == 1   # отказ запомнен
    assert report_path(root, dt.date(2026, 8, 3)).exists()

    monkeypatch.setattr(main_mod, "_today", lambda: dt.date(2027, 1, 10))
    ui = FakeUi(True)
    lines = main_mod._cleanup(root, ui)
    assert len(ui.asked) == 1 and lines[0].startswith("Удалено старых файлов: 6")
    assert not report_path(root, dt.date(2026, 10, 1)).exists()
    assert report_path(root, dt.date(2026, 11, 4)).exists()


def test_console_mode_never_deletes(root, monkeypatch):
    monkeypatch.setattr(main_mod, "_today", lambda: dt.date(2026, 11, 5))
    ui = FakeUi(True)
    ui.gui = False
    assert "без --console" in main_mod._cleanup(root, ui)[0]
    assert ui.asked == [] and not declined(root, dt.date(2026, 11, 5))
    assert report_path(root, dt.date(2026, 8, 3)).exists()
