"""
Load and validate project data from JSON or CSV files.
"""

from __future__ import annotations

import csv
import json
import pathlib

from timeline_tool.models import Milestone, Phase, Project, ReferenceLine
from timeline_tool.utils import parse_date


class LoaderError(Exception):
    """Raised when the input file is missing or malformed."""


def _parse_milestones(raw: list[dict]) -> list[Milestone]:
    return [Milestone(name=m["name"], date=parse_date(m["date"])) for m in raw]


def _parse_phases(raw: list[dict]) -> list[Phase]:
    return [
        Phase(
            name=p["name"],
            start_date=parse_date(p["start_date"]),
            end_date=parse_date(p["end_date"]),
        )
        for p in raw
    ]


def _parse_reference_lines(raw: list[dict]) -> list[ReferenceLine]:
    return [
        ReferenceLine(
            name=r["name"],
            date=parse_date(r["date"]),
            color=r.get("color", "#2196F3"),
            style=r.get("style", "-."),
        )
        for r in raw
    ]


def load_projects(filepath: str | pathlib.Path) -> tuple[list[Project], list[ReferenceLine]]:
    """
    Read *filepath* and return (projects, reference_lines).
    Supports .json and .csv files.
    """
    path = pathlib.Path(filepath)
    if not path.exists():
        raise LoaderError(f"File not found: {path}")

    if path.suffix.lower() == ".csv":
        return _load_from_csv(path), []
    else:
        return _load_from_json(path)


def _load_from_json(path: pathlib.Path) -> tuple[list[Project], list[ReferenceLine]]:
    with path.open(encoding="utf-8") as fh:
        try:
            data = json.load(fh)
        except json.JSONDecodeError as exc:
            raise LoaderError(f"Invalid JSON in {path}: {exc}") from exc

    if "projects" not in data:
        raise LoaderError("JSON must contain a top‑level 'projects' key.")

    projects: list[Project] = []
    for idx, entry in enumerate(data["projects"]):
        try:
            project = Project(
                name=entry["name"],
                start_date=parse_date(entry["start_date"]),
                end_date=parse_date(entry["end_date"]),
                color=entry.get("color"),
                status=entry.get("status", "on-track"),
                depends_on=entry.get("depends_on", []),
                phases=_parse_phases(entry.get("phases", [])),
                milestones=_parse_milestones(entry.get("milestones", [])),
            )
        except (KeyError, ValueError) as exc:
            raise LoaderError(
                f"Error in project #{idx + 1} ({entry.get('name', '?')}): {exc}"
            ) from exc

        if project.end_date <= project.start_date:
            raise LoaderError(
                f"Project '{project.name}': end_date must be after start_date."
            )
        projects.append(project)

    if not projects:
        raise LoaderError("The 'projects' list is empty.")

    ref_lines = _parse_reference_lines(data.get("reference_lines", []))

    return projects, ref_lines


def _load_from_csv(path: pathlib.Path) -> list[Project]:
    """
    Load projects from a CSV file (Enhancement #1).
    Columns: name, start_date, end_date, color (optional), status (optional)
    Milestones and phases are not supported in CSV — use JSON for full features.
    """
    projects: list[Project] = []
    with path.open(encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for idx, row in enumerate(reader):
            try:
                project = Project(
                    name=row["name"].strip(),
                    start_date=parse_date(row["start_date"].strip()),
                    end_date=parse_date(row["end_date"].strip()),
                    color=row.get("color", "").strip() or None,
                    status=row.get("status", "on-track").strip(),
                )
            except (KeyError, ValueError) as exc:
                raise LoaderError(f"Error in CSV row #{idx + 2}: {exc}") from exc

            if project.end_date <= project.start_date:
                raise LoaderError(
                    f"Project '{project.name}': end_date must be after start_date."
                )
            projects.append(project)

    if not projects:
        raise LoaderError("The CSV file has no data rows.")

    return projects