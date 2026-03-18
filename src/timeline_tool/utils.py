"""
Utility helpers — mostly date handling.
"""

from __future__ import annotations

import datetime

from timeline_tool.config import DATE_FORMAT


def parse_date(date_str: str) -> datetime.date:
    """Parse an ISO‑style date string (YYYY‑MM‑DD)."""
    return datetime.datetime.strptime(date_str, DATE_FORMAT).date()


def date_range_padded(
    projects: list,
    pad_days: int = 30,
) -> tuple[datetime.date, datetime.date]:
    """
    Return (earliest_start − pad, latest_end + pad) across all projects
    so the chart has breathing room on both sides.
    """
    starts = [p.start_date for p in projects]
    ends = [p.end_date for p in projects]
    earliest = min(starts) - datetime.timedelta(days=pad_days)
    latest = max(ends) + datetime.timedelta(days=pad_days)
    return earliest, latest