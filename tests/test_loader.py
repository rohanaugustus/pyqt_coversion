import json
import pytest
from timeline_tool.loader import load_projects, LoaderError


def test_load_sample(tmp_path):
    data = {
        "projects": [
            {
                "name": "Test",
                "start_date": "2026-01-01",
                "end_date": "2026-12-31",
            }
        ]
    }
    f = tmp_path / "test.json"
    f.write_text(json.dumps(data))
    projects = load_projects(f)
    assert len(projects) == 1
    assert projects[0].name == "Test"


def test_load_missing_file():
    with pytest.raises(LoaderError, match="File not found"):
        load_projects("/nonexistent/path.json")


def test_load_bad_dates(tmp_path):
    data = {
        "projects": [
            {
                "name": "Bad",
                "start_date": "2026-12-31",
                "end_date": "2026-01-01",
            }
        ]
    }
    f = tmp_path / "bad.json"
    f.write_text(json.dumps(data))
    with pytest.raises(LoaderError, match="end_date must be after"):
        load_projects(f)