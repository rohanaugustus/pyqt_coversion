"""
Launch the Project Timeline Tool.
- If the database exists → opens the login + editor GUI
- If not → falls back to JSON mode
"""

import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from timeline_tool.database import DEFAULT_DB_PATH

if DEFAULT_DB_PATH.exists():
    # ── Database mode: login + GUI editor ────────────────────────────
    print(f"📂 Database found at {DEFAULT_DB_PATH}")
    print("🚀 Launching GUI editor...\n")
    from timeline_tool.editor import launch
    launch()
else:
    # ── Fallback: original JSON mode ─────────────────────────────────
    print(f"⚠️  No database found at {DEFAULT_DB_PATH}")
    print("   Run 'python run_admin.py' first to set up the database.")
    print("   Falling back to JSON mode...\n")

    INPUT_FILE = os.path.join(PROJECT_ROOT, "data", "sample_projects.json")
    OUTPUT_FILE = os.path.join(PROJECT_ROOT, "output", "timeline.png")
    TITLE = "Project Timelines"
    SHOW_CHART = True

    from timeline_tool.main import main

    args = [INPUT_FILE, "--title", TITLE, "-o", OUTPUT_FILE]
    if not SHOW_CHART:
        args += ["--no-show"]

    print(f"🚀 Running timeline tool with args: {args}\n")
    main(args)