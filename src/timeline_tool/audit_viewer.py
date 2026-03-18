"""
Enhanced Audit Trail / Change History Viewer for Project Management Tool.
Provides filtering, search, export, and detailed activity viewing.
"""

from __future__ import annotations

import datetime
import csv
import pathlib
from typing import TYPE_CHECKING


def get_audit_log(
    db_path: pathlib.Path | None = None,
    limit: int = 1000,
    username: str | None = None,
    action: str | None = None,
    start_date: datetime.date | None = None,
    end_date: datetime.date | None = None,
    search_term: str | None = None,
) -> list[dict]:
    """
    Get audit log entries with optional filtering.
    
    Args:
        db_path: Database path
        limit: Maximum entries to return
        username: Filter by username
        action: Filter by action type
        start_date: Filter by start date
        end_date: Filter by end date
        search_term: Search in action or detail fields
    
    Returns:
        List of audit log entries as dicts
    """
    from timeline_tool.database import _connect
    
    query = "SELECT * FROM audit_log WHERE 1=1"
    params = []
    
    if username:
        query += " AND username = ?"
        params.append(username)
    
    if action:
        query += " AND action LIKE ?"
        params.append(f"%{action}%")
    
    if start_date:
        query += " AND DATE(timestamp) >= ?"
        params.append(start_date.strftime("%Y-%m-%d"))
    
    if end_date:
        query += " AND DATE(timestamp) <= ?"
        params.append(end_date.strftime("%Y-%m-%d"))
    
    if search_term:
        query += " AND (action LIKE ? OR detail LIKE ?)"
        params.append(f"%{search_term}%")
        params.append(f"%{search_term}%")
    
    query += " ORDER BY timestamp DESC LIMIT ?"
    params.append(limit)
    
    with _connect(db_path) as conn:
        rows = conn.execute(query, params).fetchall()
    
    return [dict(row) for row in rows]


def get_unique_users(db_path: pathlib.Path | None = None) -> list[str]:
    """Get list of unique usernames from audit log."""
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT DISTINCT username FROM audit_log ORDER BY username"
        ).fetchall()
    
    return [row["username"] for row in rows]


def get_unique_actions(db_path: pathlib.Path | None = None) -> list[str]:
    """Get list of unique action types from audit log."""
    from timeline_tool.database import _connect
    
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT DISTINCT action FROM audit_log ORDER BY action"
        ).fetchall()
    
    return [row["action"] for row in rows]


def get_activity_summary(
    db_path: pathlib.Path | None = None,
    days: int = 30,
) -> dict:
    """
    Get activity summary statistics for the past N days.
    
    Returns:
        Dict with activity statistics
    """
    from timeline_tool.database import _connect
    
    cutoff_date = (datetime.datetime.now() - datetime.timedelta(days=days)).strftime("%Y-%m-%d")
    
    with _connect(db_path) as conn:
        # Total actions
        total = conn.execute(
            "SELECT COUNT(*) as count FROM audit_log WHERE DATE(timestamp) >= ?",
            (cutoff_date,)
        ).fetchone()["count"]
        
        # Actions by type
        by_action = conn.execute(
            """SELECT action, COUNT(*) as count 
               FROM audit_log 
               WHERE DATE(timestamp) >= ?
               GROUP BY action 
               ORDER BY count DESC""",
            (cutoff_date,)
        ).fetchall()
        
        # Actions by user
        by_user = conn.execute(
            """SELECT username, COUNT(*) as count 
               FROM audit_log 
               WHERE DATE(timestamp) >= ?
               GROUP BY username 
               ORDER BY count DESC""",
            (cutoff_date,)
        ).fetchall()
        
        # Actions by day
        by_day = conn.execute(
            """SELECT DATE(timestamp) as day, COUNT(*) as count 
               FROM audit_log 
               WHERE DATE(timestamp) >= ?
               GROUP BY DATE(timestamp) 
               ORDER BY day DESC""",
            (cutoff_date,)
        ).fetchall()
    
    return {
        "total_actions": total,
        "by_action": [dict(row) for row in by_action],
        "by_user": [dict(row) for row in by_user],
        "by_day": [dict(row) for row in by_day],
        "period_days": days,
    }


def export_audit_log_csv(
    output_path: pathlib.Path,
    db_path: pathlib.Path | None = None,
    **filters,
) -> int:
    """
    Export audit log to CSV file.
    
    Args:
        output_path: Path for the CSV file
        db_path: Database path
        **filters: Filters to pass to get_audit_log()
    
    Returns:
        Number of records exported
    """
    entries = get_audit_log(db_path=db_path, limit=10000, **filters)
    
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['id', 'timestamp', 'username', 'action', 'detail'])
        writer.writeheader()
        writer.writerows(entries)
    
    return len(entries)


def get_entity_history(
    entity_type: str,
    entity_id: int,
    db_path: pathlib.Path | None = None,
) -> list[dict]:
    """
    Get change history for a specific entity (project, milestone, etc.)
    
    Args:
        entity_type: Type of entity (PROJECT, MILESTONE, PHASE, etc.)
        entity_id: ID of the entity
        db_path: Database path
    
    Returns:
        List of audit entries related to this entity
    """
    from timeline_tool.database import _connect
    
    # Build search pattern for the entity
    search_patterns = [
        f"ID {entity_id}",
        f"id {entity_id}",
        f"project {entity_id}",
        f"milestone {entity_id}",
    ]
    
    with _connect(db_path) as conn:
        results = []
        for pattern in search_patterns:
            rows = conn.execute(
                """SELECT * FROM audit_log 
                   WHERE action LIKE ? AND detail LIKE ?
                   ORDER BY timestamp DESC""",
                (f"%{entity_type}%", f"%{pattern}%")
            ).fetchall()
            results.extend([dict(row) for row in rows])
        
        # Remove duplicates (by id) while preserving order
        seen = set()
        unique_results = []
        for r in results:
            if r['id'] not in seen:
                seen.add(r['id'])
                unique_results.append(r)
        
        return sorted(unique_results, key=lambda x: x['timestamp'], reverse=True)


# Action type descriptions for display
ACTION_DESCRIPTIONS = {
    "ADD_PROJECT": "Created a new project",
    "UPDATE_PROJECT": "Updated project details",
    "DELETE_PROJECT": "Deleted a project",
    "ADD_MILESTONE": "Added a milestone",
    "UPDATE_MILESTONE": "Updated milestone details",
    "DELETE_MILESTONE": "Deleted a milestone",
    "ADD_PHASE": "Added a phase",
    "UPDATE_PHASE": "Updated phase details",
    "DELETE_PHASE": "Deleted a phase",
    "ADD_RESOURCE": "Added a team member",
    "UPDATE_RESOURCE": "Updated team member details",
    "DELETE_RESOURCE": "Removed a team member",
    "ASSIGN_RESOURCE": "Assigned team member to project",
    "REMOVE_ASSIGNMENT": "Removed team member from project",
    "CREATE_USER": "Created a user account",
    "UPDATE_ROLE": "Changed user role",
    "DELETE_USER": "Deleted a user account",
    "CHANGE_PASSWORD": "Changed password",
    "UPDATE_TASK_STATUS": "Updated task status",
    "ADD_REFLINE": "Added reference line",
    "DELETE_REFLINE": "Deleted reference line",
    "SAVE_QCTP": "Saved QCTP data",
}


def get_action_description(action: str) -> str:
    """Get human-readable description for an action type."""
    return ACTION_DESCRIPTIONS.get(action, action.replace("_", " ").title())


def get_action_icon(action: str) -> str:
    """Get an emoji icon for an action type."""
    icons = {
        "ADD": "➕",
        "CREATE": "➕",
        "UPDATE": "✏️",
        "EDIT": "✏️",
        "DELETE": "🗑️",
        "REMOVE": "🗑️",
        "ASSIGN": "🔗",
        "SAVE": "💾",
        "CHANGE": "🔄",
    }
    
    for key, icon in icons.items():
        if key in action.upper():
            return icon
    
    return "📋"