"""
Resource and team member management for Project Management Tool.
Supports team assignment, allocation tracking, and workload balancing.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from typing import TYPE_CHECKING
import pathlib
import sqlite3

if TYPE_CHECKING:
    from timeline_tool.models import Project


@dataclass
class Resource:
    """A team member or resource that can be assigned to projects."""
    
    id: int = 0
    name: str = ""
    role: str = ""  # e.g., "Developer", "Designer", "PM", "QA"
    email: str = ""
    department: str = ""
    allocation_pct: float = 100.0  # % of time available (default 100%)
    skills: list[str] = field(default_factory=list)
    
    def __str__(self):
        return f"{self.name} ({self.role})"


@dataclass
class ProjectAssignment:
    """Assignment of a resource to a project."""
    
    id: int = 0
    project_id: int = 0
    resource_id: int = 0
    role_in_project: str = ""  # Role specific to this project
    allocation_pct: float = 100.0  # % of time allocated to this project
    start_date: datetime.date | None = None
    end_date: datetime.date | None = None
    notes: str = ""


# ─────────────────────────────────────────────────────────────────────────
# Database Schema Extension
# ─────────────────────────────────────────────────────────────────────────

RESOURCE_SCHEMA = """
CREATE TABLE IF NOT EXISTS resources (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL,
    role            TEXT NOT NULL DEFAULT '',
    email           TEXT DEFAULT '',
    department      TEXT DEFAULT '',
    allocation_pct  REAL DEFAULT 100.0,
    skills          TEXT DEFAULT '',  -- JSON array as string
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS project_assignments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id      INTEGER NOT NULL,
    resource_id     INTEGER NOT NULL,
    role_in_project TEXT DEFAULT '',
    allocation_pct  REAL DEFAULT 100.0,
    start_date      TEXT,
    end_date        TEXT,
    notes           TEXT DEFAULT '',
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
    FOREIGN KEY (resource_id) REFERENCES resources(id) ON DELETE CASCADE,
    UNIQUE(project_id, resource_id)
);

CREATE INDEX IF NOT EXISTS idx_assignments_project ON project_assignments(project_id);
CREATE INDEX IF NOT EXISTS idx_assignments_resource ON project_assignments(resource_id);
"""


def init_resource_tables(db_path: pathlib.Path | None = None) -> None:
    """Initialize resource-related tables in the database."""
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        conn.executescript(RESOURCE_SCHEMA)
    print("✅ Resource tables initialized")


# ─────────────────────────────────────────────────────────────────────────
# Resource CRUD Operations
# ─────────────────────────────────────────────────────────────────────────

def add_resource(
    name: str,
    role: str = "",
    email: str = "",
    department: str = "",
    allocation_pct: float = 100.0,
    skills: list[str] | None = None,
    username: str = "system",
    db_path: pathlib.Path | None = None,
) -> int:
    """Add a new resource/team member."""
    import json
    from timeline_tool.database import _connect, log_action
    
    skills_json = json.dumps(skills or [])
    
    with _connect(db_path) as conn:
        cursor = conn.execute(
            """INSERT INTO resources (name, role, email, department, allocation_pct, skills)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (name, role, email, department, allocation_pct, skills_json)
        )
        resource_id = cursor.lastrowid
        log_action(conn, username, "ADD_RESOURCE", f"Added resource '{name}' ({role})")
    
    return resource_id


def update_resource(
    resource_id: int,
    name: str | None = None,
    role: str | None = None,
    email: str | None = None,
    department: str | None = None,
    allocation_pct: float | None = None,
    skills: list[str] | None = None,
    username: str = "system",
    db_path: pathlib.Path | None = None,
) -> None:
    """Update an existing resource."""
    import json
    from timeline_tool.database import _connect, log_action
    
    updates = []
    values = []
    
    if name is not None:
        updates.append("name = ?")
        values.append(name)
    if role is not None:
        updates.append("role = ?")
        values.append(role)
    if email is not None:
        updates.append("email = ?")
        values.append(email)
    if department is not None:
        updates.append("department = ?")
        values.append(department)
    if allocation_pct is not None:
        updates.append("allocation_pct = ?")
        values.append(allocation_pct)
    if skills is not None:
        updates.append("skills = ?")
        values.append(json.dumps(skills))
    
    if not updates:
        return
    
    updates.append("updated_at = datetime('now')")
    values.append(resource_id)
    
    with _connect(db_path) as conn:
        conn.execute(
            f"UPDATE resources SET {', '.join(updates)} WHERE id = ?",
            values
        )
        log_action(conn, username, "UPDATE_RESOURCE", f"Updated resource ID {resource_id}")


def delete_resource(
    resource_id: int,
    username: str = "system",
    db_path: pathlib.Path | None = None,
) -> None:
    """Delete a resource."""
    from timeline_tool.database import _connect, log_action
    
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM resources WHERE id = ?", (resource_id,))
        log_action(conn, username, "DELETE_RESOURCE", f"Deleted resource ID {resource_id}")


def get_all_resources(db_path: pathlib.Path | None = None) -> list[Resource]:
    """Get all resources."""
    import json
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT * FROM resources ORDER BY name"
        ).fetchall()
    
    resources = []
    for row in rows:
        skills = []
        try:
            skills = json.loads(row["skills"]) if row["skills"] else []
        except json.JSONDecodeError:
            pass
        
        resources.append(Resource(
            id=row["id"],
            name=row["name"],
            role=row["role"],
            email=row["email"] or "",
            department=row["department"] or "",
            allocation_pct=row["allocation_pct"],
            skills=skills,
        ))
    
    return resources


def get_resource_by_id(
    resource_id: int,
    db_path: pathlib.Path | None = None,
) -> Resource | None:
    """Get a specific resource by ID."""
    import json
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT * FROM resources WHERE id = ?", (resource_id,)
        ).fetchone()
    
    if not row:
        return None
    
    skills = []
    try:
        skills = json.loads(row["skills"]) if row["skills"] else []
    except json.JSONDecodeError:
        pass
    
    return Resource(
        id=row["id"],
        name=row["name"],
        role=row["role"],
        email=row["email"] or "",
        department=row["department"] or "",
        allocation_pct=row["allocation_pct"],
        skills=skills,
    )


# ─────────────────────────────────────────────────────────────────────────
# Assignment CRUD Operations
# ─────────────────────────────────────────────────────────────────────────

def assign_resource_to_project(
    project_id: int,
    resource_id: int,
    role_in_project: str = "",
    allocation_pct: float = 100.0,
    start_date: str | None = None,
    end_date: str | None = None,
    notes: str = "",
    username: str = "system",
    db_path: pathlib.Path | None = None,
) -> int:
    """Assign a resource to a project."""
    from timeline_tool.database import _connect, log_action
    
    with _connect(db_path) as conn:
        cursor = conn.execute(
            """INSERT OR REPLACE INTO project_assignments 
               (project_id, resource_id, role_in_project, allocation_pct, start_date, end_date, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (project_id, resource_id, role_in_project, allocation_pct, start_date, end_date, notes)
        )
        assignment_id = cursor.lastrowid
        log_action(conn, username, "ASSIGN_RESOURCE", 
                   f"Assigned resource {resource_id} to project {project_id}")
    
    return assignment_id


def remove_assignment(
    project_id: int,
    resource_id: int,
    username: str = "system",
    db_path: pathlib.Path | None = None,
) -> None:
    """Remove a resource assignment from a project."""
    from timeline_tool.database import _connect, log_action
    
    with _connect(db_path) as conn:
        conn.execute(
            "DELETE FROM project_assignments WHERE project_id = ? AND resource_id = ?",
            (project_id, resource_id)
        )
        log_action(conn, username, "REMOVE_ASSIGNMENT", 
                   f"Removed resource {resource_id} from project {project_id}")


def get_project_assignments(
    project_id: int,
    db_path: pathlib.Path | None = None,
) -> list[dict]:
    """Get all resource assignments for a project."""
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        rows = conn.execute(
            """SELECT pa.*, r.name, r.role, r.email, r.department
               FROM project_assignments pa
               JOIN resources r ON pa.resource_id = r.id
               WHERE pa.project_id = ?
               ORDER BY r.name""",
            (project_id,)
        ).fetchall()
    
    return [dict(row) for row in rows]


def get_resource_assignments(
    resource_id: int,
    db_path: pathlib.Path | None = None,
) -> list[dict]:
    """Get all project assignments for a resource."""
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        rows = conn.execute(
            """SELECT pa.*, p.name as project_name, p.start_date as project_start, 
                      p.end_date as project_end, p.status
               FROM project_assignments pa
               JOIN projects p ON pa.project_id = p.id
               WHERE pa.resource_id = ?
               ORDER BY p.start_date""",
            (resource_id,)
        ).fetchall()
    
    return [dict(row) for row in rows]


# ─────────────────────────────────────────────────────────────────────────
# Workload Analysis
# ─────────────────────────────────────────────────────────────────────────

def calculate_resource_utilization(
    resource_id: int,
    start_date: datetime.date | None = None,
    end_date: datetime.date | None = None,
    db_path: pathlib.Path | None = None,
) -> dict:
    """
    Calculate utilization for a resource.
    
    Returns:
        Dict with utilization metrics:
        - total_allocation: Sum of all project allocations (%)
        - available_capacity: Remaining capacity (%)
        - project_count: Number of active projects
        - status: "under", "optimal", "over"
    """
    from timeline_tool.database import _connect
    
    resource = get_resource_by_id(resource_id, db_path)
    if not resource:
        return {"error": "Resource not found"}
    
    assignments = get_resource_assignments(resource_id, db_path)
    
    # Filter by date range if provided
    today = datetime.date.today()
    start_date = start_date or today
    end_date = end_date or (today + datetime.timedelta(days=90))
    
    active_assignments = []
    for a in assignments:
        # Check if assignment overlaps with date range
        a_start = datetime.datetime.strptime(a["project_start"], "%Y-%m-%d").date()
        a_end = datetime.datetime.strptime(a["project_end"], "%Y-%m-%d").date()
        
        if a_start <= end_date and a_end >= start_date:
            active_assignments.append(a)
    
    total_allocation = sum(a["allocation_pct"] for a in active_assignments)
    max_capacity = resource.allocation_pct
    available_capacity = max_capacity - total_allocation
    
    if total_allocation > max_capacity:
        status = "over"
    elif total_allocation >= max_capacity * 0.8:
        status = "optimal"
    else:
        status = "under"
    
    return {
        "resource_id": resource_id,
        "resource_name": resource.name,
        "max_capacity": max_capacity,
        "total_allocation": total_allocation,
        "available_capacity": available_capacity,
        "project_count": len(active_assignments),
        "status": status,
        "assignments": active_assignments,
    }


def get_team_utilization_summary(
    db_path: pathlib.Path | None = None,
) -> list[dict]:
    """Get utilization summary for all resources."""
    resources = get_all_resources(db_path)
    
    summary = []
    for resource in resources:
        util = calculate_resource_utilization(resource.id, db_path=db_path)
        summary.append(util)
    
    return summary


def find_available_resources(
    allocation_needed: float = 50.0,
    role: str | None = None,
    skills: list[str] | None = None,
    db_path: pathlib.Path | None = None,
) -> list[dict]:
    """
    Find resources with available capacity.
    
    Args:
        allocation_needed: Minimum available allocation needed (%)
        role: Filter by role
        skills: Filter by required skills
    
    Returns:
        List of resources with sufficient capacity
    """
    resources = get_all_resources(db_path)
    available = []
    
    for resource in resources:
        # Check role filter
        if role and resource.role.lower() != role.lower():
            continue
        
        # Check skills filter
        if skills:
            resource_skills_lower = [s.lower() for s in resource.skills]
            if not all(s.lower() in resource_skills_lower for s in skills):
                continue
        
        # Check capacity
        util = calculate_resource_utilization(resource.id, db_path=db_path)
        if util["available_capacity"] >= allocation_needed:
            available.append({
                "resource": resource,
                "available_capacity": util["available_capacity"],
                "current_projects": util["project_count"],
            })
    
    # Sort by available capacity (descending)
    available.sort(key=lambda x: x["available_capacity"], reverse=True)
    
    return available