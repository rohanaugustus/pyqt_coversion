"""
Centralized settings — all configurable paths live here.

Resolution order (first match wins):
  1. Environment variable  PM_*
  2. Config file            ~/.project_dashboard/config.ini
  3. Built-in defaults      (local ./data directory)
"""

from __future__ import annotations

import os
import pathlib
import configparser

# ── Locate an optional config file ────────────────────────────────────
_CONFIG_DIR = pathlib.Path.home() / ".project_dashboard"
_CONFIG_FILE = _CONFIG_DIR / "config.ini"

_cfg = configparser.ConfigParser()
if _CONFIG_FILE.exists():
    _cfg.read(_CONFIG_FILE)


def _get(env_key: str, ini_section: str, ini_key: str, fallback: str) -> str:
    """Read a setting from env → config.ini → fallback, in that order."""
    # 1. Environment variable
    val = os.environ.get(env_key)
    if val:
        return val
    # 2. Config file
    if _cfg.has_option(ini_section, ini_key):
        return _cfg.get(ini_section, ini_key)
    # 3. Built-in default
    return fallback


# ── Resolved paths ────────────────────────────────────────────────────
_DEFAULT_DATA_DIR = str(pathlib.Path(__file__).resolve().parent.parent.parent / "data")

DB_DIR = pathlib.Path(
    _get("PM_DB_DIR", "database", "dir", _DEFAULT_DATA_DIR)
)
DB_PATH = DB_DIR / _get("PM_DB_NAME", "database", "filename", "timelines.db")

KPI_PATH = pathlib.Path(
    _get("PM_KPI_PATH", "kpi", "path", str(DB_DIR / "milestone_kpi.xlsx"))
)

# Logo: first check bundled resources, then allow override
_BUNDLED_LOGO = pathlib.Path(__file__).resolve().parent.parent.parent / "resources" / "icons" / "logo.png"
LOGO_PATH = pathlib.Path(
    _get("PM_LOGO_PATH", "ui", "logo_path", str(_BUNDLED_LOGO))
)