import datetime
from unittest.mock import patch
from timeline_tool.models import Project, Milestone
from timeline_tool.renderer import render_timeline


@patch("timeline_tool.renderer.plt")
def test_render_does_not_crash(mock_plt):
    """Smoke‑test: rendering completes without raising."""
    projects = [
        Project(
            name="Demo",
            start_date=datetime.date(2026, 1, 1),
            end_date=datetime.date(2026, 12, 31),
            milestones=[
                Milestone(name="Midpoint", date=datetime.date(2026, 6, 15)),
            ],
        ),
    ]
    render_timeline(projects, today=datetime.date(2026, 3, 1), show=False)