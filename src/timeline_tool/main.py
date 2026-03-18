#!/usr/bin/env python3
"""
CLI entry‑point for the timeline tool.
"""

from __future__ import annotations

import argparse
import datetime
import sys

from timeline_tool.config import DATE_FORMAT
from timeline_tool.loader import LoaderError, load_projects
from timeline_tool.renderer import render_timeline


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="timeline-tool",
        description="Visualise project timelines as a Gantt chart.",
    )
    parser.add_argument("input", help="Path to a JSON or CSV file containing project data.")
    parser.add_argument("-o", "--output", default=None, help="Save chart to file (PNG, SVG, PDF).")
    parser.add_argument("--html", default=None, help="Export interactive HTML chart (requires plotly).")
    parser.add_argument("--title", default="Project Timelines", help="Chart title.")
    parser.add_argument("--today", default=None, help="Override today's date (YYYY-MM-DD).")
    parser.add_argument("--theme", default="light", choices=["light", "dark"], help="Colour theme.")
    parser.add_argument("--no-show", action="store_true", help="Don't open Matplotlib window.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = _parse_args(argv)

    try:
        projects, ref_lines = load_projects(args.input)
    except LoaderError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"📋 Loaded {len(projects)} project(s) from {args.input}")

    today = (
        datetime.datetime.strptime(args.today, DATE_FORMAT).date()
        if args.today
        else datetime.date.today()
    )

    render_timeline(
        projects=projects,
        today=today,
        title=args.title,
        output_path=args.output,
        show=not args.no_show,
        theme=args.theme,
        reference_lines=ref_lines,
        html_path=args.html,
    )


if __name__ == "__main__":
    main()