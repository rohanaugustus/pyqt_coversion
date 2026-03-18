import datetime
import pytest
from timeline_tool.utils import parse_date, date_range_padded
from timeline_tool.models import Project


def test_parse_date_valid():
    assert parse_date("2026-03-15") == datetime.date(2026, 3, 15)


def test_parse_date_invalid():
    with pytest.raises(ValueError):
        parse_date("not-a-date")


def test_date_range_padded():
    projects = [
        Project(name="A", start_date=datetime.date(2026, 1, 1), end_date=datetime.date(2026, 6, 1)),
        Project(name="B", start_date=datetime.date(2025, 6, 1), end_date=datetime.date(2027, 1, 1)),
    ]
    lo, hi = date_range_padded(projects, pad_days=15)
    assert lo == datetime.date(2025, 6, 1) - datetime.timedelta(days=15)
    assert hi == datetime.date(2027, 1, 1) + datetime.timedelta(days=15)