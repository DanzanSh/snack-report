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
    }  # платежи (QR-код, наличные) не попадают


def test_not_an_ucs_file(tmp_path):
    bad = tmp_path / "bad.xls"
    bad.write_bytes(b"not excel")
    with pytest.raises(UcsFormatError):
        read_ucs(bad)
