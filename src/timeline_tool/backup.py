"""
Database backup and restore functionality for Project Management Tool.
Supports automatic backups, manual backup/restore, and JSON export.
"""

from __future__ import annotations

import datetime
import json
import pathlib
import shutil
import sqlite3
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from timeline_tool.models import Project


# Backup directory (relative to database location)
def get_backup_dir(db_path: pathlib.Path) -> pathlib.Path:
    """Get the backup directory for a given database path."""
    backup_dir = db_path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    return backup_dir


def create_backup(
    db_path: pathlib.Path,
    backup_name: str | None = None,
    max_backups: int = 10,
) -> pathlib.Path:
    """
    Create a timestamped backup of the database.
    
    Args:
        db_path: Path to the database file
        backup_name: Optional custom name for the backup
        max_backups: Maximum number of backups to keep (oldest are deleted)
    
    Returns:
        Path to the created backup file
    """
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found: {db_path}")
    
    backup_dir = get_backup_dir(db_path)
    
    # Generate backup filename
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if backup_name:
        backup_filename = f"backup_{backup_name}_{timestamp}.db"
    else:
        backup_filename = f"backup_{timestamp}.db"
    
    backup_path = backup_dir / backup_filename
    
    # Copy the database file
    shutil.copy2(db_path, backup_path)
    
    # Clean up old backups if needed
    _cleanup_old_backups(backup_dir, max_backups)
    
    print(f"✅ Backup created: {backup_path}")
    return backup_path


def restore_backup(
    backup_path: pathlib.Path,
    db_path: pathlib.Path,
    create_pre_restore_backup: bool = True,
) -> bool:
    """
    Restore the database from a backup.
    
    Args:
        backup_path: Path to the backup file
        db_path: Path to the target database file
        create_pre_restore_backup: Whether to create a backup before restoring
    
    Returns:
        True if successful
    """
    if not backup_path.exists():
        raise FileNotFoundError(f"Backup not found: {backup_path}")
    
    # Create a pre-restore backup
    if create_pre_restore_backup and db_path.exists():
        create_backup(db_path, backup_name="pre_restore")
    
    # Restore by copying
    shutil.copy2(backup_path, db_path)
    
    print(f"✅ Database restored from: {backup_path}")
    return True


def list_backups(db_path: pathlib.Path) -> list[dict]:
    """
    List all available backups.
    
    Returns:
        List of dicts with backup info: {path, name, created, size_kb}
    """
    backup_dir = get_backup_dir(db_path)
    backups = []
    
    for file in sorted(backup_dir.glob("backup_*.db"), reverse=True):
        stat = file.stat()
        backups.append({
            "path": file,
            "name": file.stem,
            "created": datetime.datetime.fromtimestamp(stat.st_mtime),
            "size_kb": stat.st_size / 1024,
        })
    
    return backups


def delete_backup(backup_path: pathlib.Path) -> bool:
    """Delete a specific backup file."""
    if backup_path.exists():
        backup_path.unlink()
        print(f"🗑️ Backup deleted: {backup_path}")
        return True
    return False


def _cleanup_old_backups(backup_dir: pathlib.Path, max_backups: int) -> int:
    """Remove oldest backups if there are more than max_backups."""
    backups = sorted(backup_dir.glob("backup_*.db"), key=lambda p: p.stat().st_mtime)
    removed = 0
    
    while len(backups) > max_backups:
        oldest = backups.pop(0)
        oldest.unlink()
        removed += 1
    
    return removed


def export_to_json(
    db_path: pathlib.Path,
    output_path: pathlib.Path,
) -> bool:
    """
    Export the entire database to a portable JSON file.
    
    Args:
        db_path: Path to the database file
        output_path: Path for the JSON output file
    
    Returns:
        True if successful
    """
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    
    try:
        data = {
            "exported_at": datetime.datetime.now().isoformat(),
            "version": "1.0",
            "projects": [],
            "milestones": [],
            "phases": [],
            "milestone_tasks": [],
            "reference_lines": [],
            "qctp": [],
            "users": [],  # Note: passwords are hashed, but we include usernames/roles
        }
        
        # Export projects
        rows = conn.execute("SELECT * FROM projects").fetchall()
        data["projects"] = [dict(row) for row in rows]
        
        # Export milestones
        rows = conn.execute("SELECT * FROM milestones").fetchall()
        data["milestones"] = [dict(row) for row in rows]
        
        # Export phases
        rows = conn.execute("SELECT * FROM phases").fetchall()
        data["phases"] = [dict(row) for row in rows]
        
        # Export milestone_tasks
        rows = conn.execute("SELECT * FROM milestone_tasks").fetchall()
        data["milestone_tasks"] = [dict(row) for row in rows]
        
        # Export reference_lines
        rows = conn.execute("SELECT * FROM reference_lines").fetchall()
        data["reference_lines"] = [dict(row) for row in rows]
        
        # Export QCTP
        rows = conn.execute("SELECT * FROM qctp").fetchall()
        data["qctp"] = [dict(row) for row in rows]
        
        # Export users (excluding passwords)
        rows = conn.execute("SELECT username, role, full_name, created_at FROM users").fetchall()
        data["users"] = [dict(row) for row in rows]
        
        # Write JSON
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, default=str)
        
        print(f"✅ Database exported to JSON: {output_path}")
        return True
        
    finally:
        conn.close()


def import_from_json(
    json_path: pathlib.Path,
    db_path: pathlib.Path,
    merge: bool = False,
) -> dict:
    """
    Import data from a JSON file into the database.
    
    Args:
        json_path: Path to the JSON file
        db_path: Path to the database file
        merge: If True, merge with existing data; if False, replace
    
    Returns:
        Dict with import statistics
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    conn = sqlite3.connect(str(db_path))
    stats = {"projects": 0, "milestones": 0, "phases": 0, "tasks": 0}
    
    try:
        if not merge:
            # Clear existing data (except users)
            conn.execute("DELETE FROM milestone_tasks")
            conn.execute("DELETE FROM milestones")
            conn.execute("DELETE FROM phases")
            conn.execute("DELETE FROM qctp")
            conn.execute("DELETE FROM projects")
            conn.execute("DELETE FROM reference_lines")
        
        # Import projects
        for proj in data.get("projects", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO projects 
                       (id, name, start_date, end_date, color, status, dev_region, sales_region, created_by, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (proj.get("id"), proj["name"], proj["start_date"], proj["end_date"],
                     proj.get("color"), proj.get("status", "on-track"),
                     proj.get("dev_region", ""), proj.get("sales_region", ""),
                     proj.get("created_by"), proj.get("updated_at"))
                )
                stats["projects"] += 1
            except sqlite3.IntegrityError:
                pass
        
        # Import milestones
        for ms in data.get("milestones", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO milestones 
                       (id, project_id, name, date, created_by, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (ms.get("id"), ms["project_id"], ms["name"], ms["date"],
                     ms.get("created_by"), ms.get("updated_at"))
                )
                stats["milestones"] += 1
            except sqlite3.IntegrityError:
                pass
        
        # Import phases
        for ph in data.get("phases", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO phases 
                       (id, project_id, name, start_date, end_date, created_by, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (ph.get("id"), ph["project_id"], ph["name"], ph["start_date"],
                     ph["end_date"], ph.get("created_by"), ph.get("updated_at"))
                )
                stats["phases"] += 1
            except sqlite3.IntegrityError:
                pass
        
        # Import milestone tasks
        for task in data.get("milestone_tasks", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO milestone_tasks 
                       (id, milestone_id, task_name, status, attachment_path, updated_by, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (task.get("id"), task["milestone_id"], task["task_name"],
                     task.get("status", "Yet to Start"), task.get("attachment_path"),
                     task.get("updated_by"), task.get("updated_at"))
                )
                stats["tasks"] += 1
            except sqlite3.IntegrityError:
                pass
        
        # Import reference lines
        for ref in data.get("reference_lines", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO reference_lines 
                       (id, name, date, color, style)
                       VALUES (?, ?, ?, ?, ?)""",
                    (ref.get("id"), ref["name"], ref["date"],
                     ref.get("color", "#2196F3"), ref.get("style", "-."))
                )
            except sqlite3.IntegrityError:
                pass
        
        # Import QCTP
        for qctp in data.get("qctp", []):
            try:
                conn.execute(
                    """INSERT OR REPLACE INTO qctp 
                       (id, project_id, quality, cost, time, performance, updated_by, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (qctp.get("id"), qctp["project_id"], qctp.get("quality", ""),
                     qctp.get("cost", ""), qctp.get("time", ""), qctp.get("performance", ""),
                     qctp.get("updated_by"), qctp.get("updated_at"))
                )
            except sqlite3.IntegrityError:
                pass
        
        conn.commit()
        print(f"✅ Import complete: {stats}")
        return stats
        
    finally:
        conn.close()


def auto_backup_before_migration(db_path: pathlib.Path) -> pathlib.Path | None:
    """
    Create an automatic backup before database schema changes.
    Called by migrate_db() function.
    """
    if db_path.exists():
        return create_backup(db_path, backup_name="pre_migration")
    return None