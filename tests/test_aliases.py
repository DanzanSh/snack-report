"""Сопоставления названий UCS со строками отчёта из config/settings.ini."""
import datetime as dt
import shutil

import openpyxl
import pytest

from conftest import ROOT
from snack_report.cash_report import get_snack_sheet
from snack_report.main import load_settings
from snack_report.report_builder import DEFAULT_ALIASES, build_report, normalize

SETTINGS = load_settings(ROOT)


def _template_rows(path=ROOT / "config" / "template.xlsx") -> dict[str, str]:
    ws = get_snack_sheet(openpyxl.load_workbook(path))
    return {normalize(v): v for r in range(3, ws.max_row + 1) if (v := ws.cell(r, 1).value)}


def _ini_aliases() -> dict[str, str]:
    """Только то, что записано в config/settings.ini, без встроенных."""
    import configparser
    cfg = configparser.ConfigParser(delimiters=("=",))
    cfg.optionxform = str
    cfg.read(ROOT / "config" / "settings.ini", encoding="utf-8-sig")
    return {k.strip(): v.strip() for k, v in cfg.items("aliases")}


def test_every_alias_target_is_a_template_row():
    rows = _template_rows()
    missing = {k: v for k, v in SETTINGS["aliases"].items() if normalize(v) not in rows}
    assert missing == {}


def test_every_price_split_target_is_a_template_row():
    rows = _template_rows()
    targets = [row for prices in SETTINGS["price_split"].values() for row in prices.values()]
    assert [t for t in targets if normalize(t) not in rows] == []


def test_settings_ini_repeats_builtin_aliases():
    """В settings.ini сказано, что сопоставления уже встроены: обе копии должны совпадать
    (так ловится расхождение вроде «Левиной» / «Леврон»)."""
    ini = {normalize(k): normalize(v) for k, v in _ini_aliases().items()}
    builtin = {normalize(k): normalize(v) for k, v in DEFAULT_ALIASES.items()}
    assert {k: v for k, v in builtin.items() if ini.get(k) != v} == {}
    # в ini может быть больше, но только сопоставления, которые ничего не меняют («Пепси Кола» = «Пепси кола»)
    extra = {k: v for k, v in ini.items() if k not in builtin and k != v}
    assert extra == {}


@pytest.mark.parametrize("ucs_name, row", [
    ("Дада", "Сок Дада пластик бут 0.5"),
    ("Витамин D", "Витамин D3 180 капсул"),
    ("Пепси Кола", "Пепси кола"),
    ("Бомбар Молочный Коктейль", "Бомбар молочный коктейль"),
    ("Шот гуарана", "Шот Гуарана"),
    ("Протеин Леврон Черный 2кг", "Протеин оптом"),
    ("Креатин порция", "Креатин порц"),
    ("Вода 0.5", "Вода 0,5"),
])
def test_alias_sale_goes_to_existing_row(tmp_path, ucs_name, row):
    tpl = tmp_path / "template.xlsx"
    shutil.copy(ROOT / "config" / "template.xlsx", tpl)
    out = tmp_path / "out.xlsx"
    res = build_report(date=dt.date(2026, 11, 5), sales={ucs_name: 2}, source=tpl, from_template=True,
                       output=out, section_rows=SETTINGS["section_rows"], staff_rows=SETTINGS["staff_rows"],
                       aliases=SETTINGS["aliases"], ignored=SETTINGS["ignored"])
    assert res.added == []  # не появилось новой строки-дубля
    ws = get_snack_sheet(openpyxl.load_workbook(out))
    r = next(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value == row)
    assert ws.cell(r, 4).value == 2


def test_builtin_aliases_work_without_settings_ini(tmp_path):
    """exe без settings.ini: новые сопоставления тоже встроены."""
    s = load_settings(tmp_path)
    assert s["aliases"]["Дада"] == "Сок Дада пластик бут 0.5"
    assert s["aliases"]["Витамин D"] == "Витамин D3 180 капсул"
    assert s["aliases"]["Протеин Леврон Черный 2кг"] == "Протеин оптом"
