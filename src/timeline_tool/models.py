"""
Data models for projects, milestones, phases, and reference lines.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field

@dataclass
class Project:
    """A project with a date range, status, phases, milestones, and dependencies."""

    name: str
    start_date: datetime.date
    end_date: datetime.date
    color: str | None = None
    status: str = "on-track"          # on-track | at-risk | overdue (legacy/default)
    dev_region: str = ""
    sales_region: str = ""
    depends_on: list[str] = field(default_factory=list)
    phases: list[Phase] = field(default_factory=list)
    milestones: list[Milestone] = field(default_factory=list)

    def duration_days(self) -> int:
        return (self.end_date - self.start_date).days

    def progress(self, today: datetime.date | None = None) -> float:
        today = today or datetime.date.today()
        if today <= self.start_date:
            return 0.0
        if today >= self.end_date:
            return 1.0
        elapsed = (today - self.start_date).days
        total = self.duration_days()
        return elapsed / total if total > 0 else 0.0

    def computed_status(self, today: datetime.date | None = None) -> str:
        """
        Automatically compute project status from milestone task completion.

        Rules (evaluated per milestone, worst status wins):
        - OVERDUE:   today > milestone.date AND milestone has incomplete tasks
        - AT RISK:   milestone.date <= today + 5 days AND milestone has incomplete tasks
        - ON TRACK:  all tasks completed at least 5 days before milestone date, 
                     or no pending milestones

        Returns: "overdue", "at-risk", or "on-track"
        """
        today = today or datetime.date.today()

        if not self.milestones:
            return "on-track"

        worst = "on-track"  # best case

        for ms in self.milestones:
            # Check if milestone has incomplete tasks
            has_tasks = bool(ms.task_statuses) and sum(ms.task_statuses.values()) > 0
            if not has_tasks:
                continue  # No tasks defined → skip this milestone

            if ms.is_complete():
                continue  # All tasks done → this milestone is fine

            # Milestone has incomplete tasks — evaluate against dates
            days_until = (ms.date - today).days

            if days_until < 0:
                # Today has crossed the milestone date with pending tasks
                return "overdue"  # Worst possible — return immediately
            elif days_until <= 5:
                # Within 5 days of milestone and tasks still pending
                worst = "at-risk"
            # else: more than 5 days away — still on-track for this milestone

        return worst

@dataclass
class Milestone:
    """A named point‑in‑time event within a project."""
    name: str
    date: datetime.date
    tasks: list[str] = field(default_factory=list)
    milestone_id: int = 0
    task_statuses: dict = field(default_factory=dict)  # {"Completed": 2, "WIP": 1, ...}
    
    def is_complete(self) -> bool:
        """Check if all tasks are completed."""
        if not self.task_statuses:
            return False
        total_tasks = sum(self.task_statuses.values())
        if total_tasks == 0:
            return False
        completed = self.task_statuses.get("Completed", 0)
        not_applicable = self.task_statuses.get("Not Applicable", 0)
        # Consider complete if all tasks are either Completed or Not Applicable
        return (completed + not_applicable) == total_tasks

    def marker_color(self, today: datetime.date | None = None) -> str:
        """
        Determine the milestone marker color based on status and proximity.

        Rules:
        - Green (#27AE60):    All tasks are Completed (or Completed + Not Applicable).
        - Orange (#F39C12):   Less than 5 days to milestone date AND has
                              'Yet to Start' or 'WIP' tasks still pending.
        - Red (#E74C3C):      Milestone date has passed AND has
                              'Yet to Start' or 'WIP' tasks still pending.
        - Light Blue (#5DADE2): Upcoming milestone with more than 5 days
                                from today (not yet at risk).

        Returns: hex color string
        """
        today = today or datetime.date.today()

        # ── Completed → Green ────────────────────────────────────────
        if self.is_complete():
            return "#27AE60"

        # Check for incomplete work (Yet to Start or WIP)
        has_tasks = bool(self.task_statuses) and sum(self.task_statuses.values()) > 0
        if not has_tasks:
            # No tasks defined at all → treat as upcoming (light blue)
            return "#5DADE2"

        yet_to_start = self.task_statuses.get("Yet to Start", 0)
        wip = self.task_statuses.get("WIP", 0)
        has_pending = (yet_to_start + wip) > 0

        if not has_pending:
            # Edge case: tasks exist but none are YTS/WIP
            # (e.g. all "Not Applicable" but is_complete() was False
            #  because there's some other status)
            return "#5DADE2"

        days_until = (self.date - today).days

        # ── Overdue → Red ────────────────────────────────────────────
        if days_until < 0:
            return "#E74C3C"

        # ── At Risk → Orange ─────────────────────────────────────────
        if days_until <= 5:
            return "#F39C12"

        # ── Upcoming → Light Blue ────────────────────────────────────
        return "#5DADE2"

@dataclass(frozen=True)
class Phase:
    """A sub‑period within a project (e.g. Design, Development, Testing)."""
    name: str
    start_date: datetime.date
    end_date: datetime.date


@dataclass(frozen=True)
class ReferenceLine:
    """A custom vertical reference line (e.g. Board Meeting, Budget Review)."""
    name: str
    date: datetime.date
    color: str = "#2196F3"
    style: str = "-."