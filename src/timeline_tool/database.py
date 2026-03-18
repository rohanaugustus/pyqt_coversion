"""
SQLite database layer for projects, milestones, phases, and users.
Stores the DB at a shared network path so multiple users can access it.
"""

from __future__ import annotations

import datetime
import sqlite3
import pathlib
import os
import time  # ✅ ADD THIS
import random  # ✅ ADD THIS
from contextlib import contextmanager

from timeline_tool.models import Project, Milestone, Phase, ReferenceLine
from timeline_tool.utils import parse_date
from timeline_tool.resources import RESOURCE_SCHEMA

# ── Default DB path (shared network folder) - Dynamic per user ─────────
current_user = os.environ.get('USERNAME', 'T0276HS')
DEFAULT_DB_DIR = pathlib.Path(
    f"C:/Users/{current_user}/Stellantis/AI ML - VEHE STRUCTURE - Documents/Shared/project_timelines"
)
DEFAULT_DB_PATH = DEFAULT_DB_DIR / "timelines.db"

# ✅ ADD THESE CONSTANTS FOR MULTI-USER SUPPORT ────────────────────────
MAX_RETRIES = 5
RETRY_DELAY_MIN = 0.1  # seconds
RETRY_DELAY_MAX = 1.0  # seconds
CONNECTION_TIMEOUT = 60  # seconds


# ✅ REPLACE THE _connect FUNCTION WITH THIS ───────────────────────────
@contextmanager
def _connect(db_path: pathlib.Path | None = None):
    """
    Context manager for a SQLite connection with enhanced multi-user support.
    Includes retry logic for database locks and proper transaction handling.
    """
    path = db_path or DEFAULT_DB_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    
    # Ensure we're using a consistent path (resolve any symlinks)
    path = path.resolve()
    
    conn = None
    retry_count = 0
    last_error = None
    
    while retry_count < MAX_RETRIES:
        try:
            # Create connection with longer timeout
            conn = sqlite3.connect(str(path), timeout=CONNECTION_TIMEOUT, check_same_thread=False)
            
            # Configure for better concurrent access on network drives
            conn.execute("PRAGMA journal_mode=DELETE;")  # More reliable on network drives than WAL
            conn.execute("PRAGMA synchronous=FULL;")     # Ensure data integrity
            conn.execute("PRAGMA foreign_keys=ON;")
            conn.execute("PRAGMA busy_timeout=30000;")   # 30 second busy timeout
            conn.execute("PRAGMA temp_store=MEMORY;")    # Use memory for temp storage
            conn.execute("PRAGMA cache_size=-10000;")    # 10MB cache
            
            conn.row_factory = sqlite3.Row
            
            # Test the connection with a simple query
            conn.execute("SELECT 1").fetchone()
            
            break  # Connection successful
            
        except sqlite3.OperationalError as e:
            last_error = e
            retry_count += 1
            
            if retry_count < MAX_RETRIES:
                # Random backoff to avoid thundering herd
                delay = random.uniform(RETRY_DELAY_MIN, RETRY_DELAY_MAX)
                print(f"⚠️  Database locked, retrying in {delay:.2f}s... (attempt {retry_count}/{MAX_RETRIES})")
                time.sleep(delay)
                
                if conn:
                    try:
                        conn.close()
                    except:
                        pass
                    conn = None
            else:
                if conn:
                    try:
                        conn.close()
                    except:
                        pass
                raise Exception(f"Failed to connect to database after {MAX_RETRIES} attempts: {last_error}")
    
    if conn is None:
        raise Exception(f"Failed to establish database connection: {last_error}")
    
    try:
        yield conn
        conn.commit()
    except sqlite3.OperationalError as e:
        # Handle database locked errors during transaction
        if "locked" in str(e).lower():
            print(f"⚠️  Database locked during transaction, rolling back...")
            try:
                conn.rollback()
            except:
                pass
            raise
        else:
            try:
                conn.rollback()
            except:
                pass
            raise
    except Exception as e:
        try:
            conn.rollback()
        except:
            pass
        raise
    finally:
        try:
            conn.close()
        except:
            pass


# ✅ ADD THIS NEW FUNCTION ──────────────────────────────────────────────
def execute_with_retry(func, *args, max_retries=3, **kwargs):
    """
    Execute a database function with retry logic for handling locks.
    """
    retry_count = 0
    last_error = None
    
    while retry_count < max_retries:
        try:
            return func(*args, **kwargs)
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower():
                retry_count += 1
                last_error = e
                
                if retry_count < max_retries:
                    delay = random.uniform(RETRY_DELAY_MIN, RETRY_DELAY_MAX)
                    print(f"⚠️  Operation failed due to lock, retrying in {delay:.2f}s... (attempt {retry_count}/{max_retries})")
                    time.sleep(delay)
                else:
                    raise Exception(f"Operation failed after {max_retries} attempts: {last_error}")
            else:
                raise
        except Exception as e:
            raise
    
    if last_error:
        raise last_error


# ─────────────────────────────────────────────────────────────────────────
# Schema creation
# ─────────────────────────────────────────────────────────────────────────

# ✅ REPLACE init_db WITH THIS ──────────────────────────────────────────
def init_db(db_path: pathlib.Path | None = None) -> None:
    """Create all tables if they don't exist."""
    def _create_tables(db_path):
        with _connect(db_path) as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS users (
                    username    TEXT PRIMARY KEY,
                    password    TEXT NOT NULL,
                    role        TEXT NOT NULL DEFAULT 'viewer',
                    full_name   TEXT,
                    created_at  TEXT NOT NULL DEFAULT (datetime('now'))
                );

                CREATE TABLE IF NOT EXISTS projects (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    name        TEXT NOT NULL UNIQUE,
                    start_date  TEXT NOT NULL,
                    end_date    TEXT NOT NULL,
                    color       TEXT,
                    status      TEXT NOT NULL DEFAULT 'on-track',
                    dev_region  TEXT NOT NULL DEFAULT '',
                    sales_region TEXT NOT NULL DEFAULT '',
                    created_by  TEXT,
                    updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
                );

                CREATE TABLE IF NOT EXISTS milestones (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id  INTEGER NOT NULL,
                    name        TEXT NOT NULL,
                    date        TEXT NOT NULL,
                    created_by  TEXT,
                    updated_at  TEXT NOT NULL DEFAULT (datetime('now')),
                    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS milestone_tasks (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    milestone_id    INTEGER NOT NULL,
                    task_name       TEXT NOT NULL,
                    status          TEXT NOT NULL DEFAULT 'Yet to Start',
                    attachment_path TEXT,
                    updated_by      TEXT,
                    updated_at      TEXT NOT NULL DEFAULT (datetime('now')),
                    FOREIGN KEY (milestone_id) REFERENCES milestones(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS phases (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id  INTEGER NOT NULL,
                    name        TEXT NOT NULL,
                    start_date  TEXT NOT NULL,
                    end_date    TEXT NOT NULL,
                    created_by  TEXT,
                    updated_at  TEXT NOT NULL DEFAULT (datetime('now')),
                    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS reference_lines (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    name        TEXT NOT NULL,
                    date        TEXT NOT NULL,
                    color       TEXT DEFAULT '#2196F3',
                    style       TEXT DEFAULT '-.'
                );

                CREATE TABLE IF NOT EXISTS audit_log (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    username    TEXT NOT NULL,
                    action      TEXT NOT NULL,
                    detail      TEXT,
                    timestamp   TEXT NOT NULL DEFAULT (datetime('now'))
                );
                
                CREATE TABLE IF NOT EXISTS qctp (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id  INTEGER NOT NULL UNIQUE,
                    quality     TEXT DEFAULT '',
                    cost        TEXT DEFAULT '',
                    time        TEXT DEFAULT '',
                    performance TEXT DEFAULT '',
                    updated_by  TEXT,
                    updated_at  TEXT NOT NULL DEFAULT (datetime('now')),
                    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS activities (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id      INTEGER NOT NULL,
                    week_number     INTEGER NOT NULL,
                    year            INTEGER NOT NULL,
                    activity_name   TEXT NOT NULL,
                    start_date      TEXT NOT NULL,
                    end_date        TEXT NOT NULL,
                    time_taken      TEXT DEFAULT '',
                    members         TEXT DEFAULT '',
                    hard_points     TEXT DEFAULT '',
                    status          TEXT NOT NULL DEFAULT 'WIP',
                    attachment_path TEXT DEFAULT '',
                    created_by      TEXT,
                    updated_by      TEXT,
                    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
                    updated_at      TEXT NOT NULL DEFAULT (datetime('now')),
                    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_activities_project ON activities(project_id);
                CREATE INDEX IF NOT EXISTS idx_activities_week ON activities(project_id, year, week_number);
                CREATE INDEX IF NOT EXISTS idx_milestone_tasks_milestone ON milestone_tasks(milestone_id);
                CREATE INDEX IF NOT EXISTS idx_phases_project ON phases(project_id);
                CREATE INDEX IF NOT EXISTS idx_qctp_project ON qctp(project_id);
            """)
        print(f"✅ Database initialized at {db_path or DEFAULT_DB_PATH}")
    
    execute_with_retry(_create_tables, db_path)


def init_resource_tables(db_path: pathlib.Path | None = None) -> None:
    """Initialize resource-related tables in the database."""
    with _connect(db_path) as conn:
        conn.executescript(RESOURCE_SCHEMA)
    print("✅ Resource tables initialized")

def migrate_db(db_path: pathlib.Path | None = None) -> None:
    """Run all database migrations."""
    with _connect(db_path) as conn:
        # ... existing migration code ...
        
        # Add QCTP line items table
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS activities (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id      INTEGER NOT NULL,
                week_number     INTEGER NOT NULL,
                year            INTEGER NOT NULL,
                activity_name   TEXT NOT NULL,
                start_date      TEXT NOT NULL,
                end_date        TEXT NOT NULL,
                time_taken      TEXT DEFAULT '',
                members         TEXT DEFAULT '',
                hard_points     TEXT DEFAULT '',
                status          TEXT NOT NULL DEFAULT 'WIP',
                attachment_path TEXT DEFAULT '',
                created_by      TEXT,
                updated_by      TEXT,
                created_at      TEXT NOT NULL DEFAULT (datetime('now')),
                updated_at      TEXT NOT NULL DEFAULT (datetime('now')),
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_activities_project ON activities(project_id);
            CREATE INDEX IF NOT EXISTS idx_activities_week ON activities(project_id, year, week_number);
        """)
        
        print("✅ Database migration complete (QCTP line items table ready)")

# ✅ UPDATE log_action TO NOT FAIL ON ERROR ─────────────────────────────
def log_action(conn: sqlite3.Connection, username: str, action: str, detail: str = "") -> None:
    """Log an action to the audit log."""
    try:
        conn.execute(
            "INSERT INTO audit_log (username, action, detail) VALUES (?, ?, ?)",
            (username, action, detail),
        )
    except sqlite3.OperationalError as e:
        # Don't fail the whole operation if audit log fails
        print(f"⚠️  Warning: Failed to write audit log: {e}")


# ─────────────────────────────────────────────────────────────────────────
# Load all projects with milestones and phases
# ─────────────────────────────────────────────────────────────────────────

def load_all(db_path: pathlib.Path | None = None) -> tuple[list[Project], list[ReferenceLine]]:
    """Load all projects, milestones, phases, and reference lines from the database."""
    
    # Sync KPI tasks to database first
    sync_milestone_tasks_from_kpi(db_path)
    
    with _connect(db_path) as conn:
        # Projects
        proj_rows = conn.execute("SELECT * FROM projects ORDER BY start_date").fetchall()
        projects: list[Project] = []
        
        for proj in proj_rows:
            # Milestones for this project
            ms_rows = conn.execute(
                "SELECT id, name, date FROM milestones WHERE project_id = ? ORDER BY date",
                (proj["id"],)
            ).fetchall()
            
            milestones = []
            for ms in ms_rows:
                ms_id = ms["id"]
                ms_name = ms["name"]
                
                # Get tasks from database with status
                task_rows = conn.execute(
                    "SELECT task_name, status, attachment_path FROM milestone_tasks WHERE milestone_id = ?",
                    (ms_id,)
                ).fetchall()
                
                tasks = [row["task_name"] for row in task_rows]
                
                # Calculate status summary
                task_statuses = {
                    "Completed": 0,
                    "WIP": 0,
                    "Yet to Start": 0,
                    "Not Applicable": 0
                }
                
                for row in task_rows:
                    status = row["status"]
                    if status in task_statuses:
                        task_statuses[status] += 1
                
                milestone = Milestone(
                    name=ms_name,
                    date=parse_date(ms["date"]),
                    tasks=tasks,
                    milestone_id=ms_id,
                    task_statuses=task_statuses
                )
                milestones.append(milestone)
            
            # Phases for this project
            ph_rows = conn.execute(
                "SELECT name, start_date, end_date FROM phases WHERE project_id = ? ORDER BY start_date",
                (proj["id"],)
            ).fetchall()
            phases = [
                Phase(
                    name=ph["name"],
                    start_date=parse_date(ph["start_date"]),
                    end_date=parse_date(ph["end_date"]),
                )
                for ph in ph_rows
            ]
            
            # Build Project
            project = Project(
                name=proj["name"],
                start_date=parse_date(proj["start_date"]),
                end_date=parse_date(proj["end_date"]),
                color=proj["color"],
                status=proj["status"],
                dev_region=proj["dev_region"] if "dev_region" in proj.keys() else "",
                sales_region=proj["sales_region"] if "sales_region" in proj.keys() else "",
                milestones=milestones,
                phases=phases,
            )
            projects.append(project)
        
        # Reference lines
        ref_rows = conn.execute("SELECT * FROM reference_lines ORDER BY date").fetchall()
        ref_lines = [
            ReferenceLine(
                name=r["name"],
                date=parse_date(r["date"]),
                color=r["color"],
                style=r["style"],
            )
            for r in ref_rows
        ]
    
    return projects, ref_lines

# ─────────────────────────────────────────────────────────────────────────
# CRUD operations - Keep all existing functions as-is
# ─────────────────────────────────────────────────────────────────────────

def add_project(name: str, start_date: str, end_date: str, color: str | None = None,
                status: str = "on-track", dev_region: str = "", sales_region: str = "",
                username: str = "system",
                db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO projects (name, start_date, end_date, color, status, dev_region, sales_region, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (name, start_date, end_date, color, status, dev_region, sales_region, username),
        )
        log_action(conn, username, "ADD_PROJECT", f"Added project '{name}'")


def update_project(project_id: int, name: str, start_date: str, end_date: str,
                   color: str | None = None, status: str = "on-track",
                   dev_region: str = "", sales_region: str = "",
                   username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE projects SET name=?, start_date=?, end_date=?, color=?, status=?, dev_region=?, sales_region=?, updated_at=datetime('now') WHERE id=?",
            (name, start_date, end_date, color, status, dev_region, sales_region, project_id),
        )
        log_action(conn, username, "UPDATE_PROJECT", f"Updated project ID {project_id}")

def delete_project(project_id: int, username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
        log_action(conn, username, "DELETE_PROJECT", f"Deleted project ID {project_id}")


def add_milestone(project_id: int, name: str, date: str, username: str = "system",
                  db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO milestones (project_id, name, date, created_by) VALUES (?, ?, ?, ?)",
            (project_id, name, date, username),
        )
        log_action(conn, username, "ADD_MILESTONE", f"Added milestone '{name}' to project {project_id}")


def update_milestone(milestone_id: int, name: str, date: str, username: str = "system",
                     db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE milestones SET name=?, date=?, updated_at=datetime('now') WHERE id=?",
            (name, date, milestone_id),
        )
        log_action(conn, username, "UPDATE_MILESTONE", f"Updated milestone ID {milestone_id}")


def delete_milestone(milestone_id: int, username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM milestones WHERE id = ?", (milestone_id,))
        log_action(conn, username, "DELETE_MILESTONE", f"Deleted milestone ID {milestone_id}")


def add_phase(project_id: int, name: str, start_date: str, end_date: str,
              username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO phases (project_id, name, start_date, end_date, created_by) VALUES (?, ?, ?, ?, ?)",
            (project_id, name, start_date, end_date, username),
        )
        log_action(conn, username, "ADD_PHASE", f"Added phase '{name}' to project {project_id}")

def update_phase(phase_id: int, name: str, start_date: str, end_date: str,
                 username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE phases SET name = ?, start_date = ?, end_date = ?, created_by = ?, "
            "updated_at = datetime('now') WHERE id = ?",
            (name, start_date, end_date, username, phase_id),
        )
        log_action(conn, username, "UPDATE_PHASE", f"Updated phase ID {phase_id}: '{name}'")

def delete_phase(phase_id: int, username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM phases WHERE id = ?", (phase_id,))
        log_action(conn, username, "DELETE_PHASE", f"Deleted phase ID {phase_id}")


def add_reference_line(name: str, date: str, color: str = "#2196F3", style: str = "-.",
                       username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO reference_lines (name, date, color, style) VALUES (?, ?, ?, ?)",
            (name, date, color, style),
        )
        log_action(conn, username, "ADD_REFLINE", f"Added reference line '{name}'")


def delete_reference_line(ref_id: int, username: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM reference_lines WHERE id = ?", (ref_id,))
        log_action(conn, username, "DELETE_REFLINE", f"Deleted reference line ID {ref_id}")


def import_from_json(json_path: str | pathlib.Path, username: str = "system",
                     db_path: pathlib.Path | None = None) -> None:
    """Import projects from a JSON file."""
    import json
    from pathlib import Path
    
    path = Path(json_path)
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    
    for proj_data in data.get("projects", []):
        with _connect(db_path) as conn:
            # Add project
            cursor = conn.execute(
                "INSERT INTO projects (name, start_date, end_date, color, status, created_by) VALUES (?, ?, ?, ?, ?, ?)",
                (proj_data["name"], proj_data["start_date"], proj_data["end_date"],
                 proj_data.get("color"), proj_data.get("status", "on-track"), username),
            )
            proj_id = cursor.lastrowid
            
            # Add milestones
            for ms in proj_data.get("milestones", []):
                conn.execute(
                    "INSERT INTO milestones (project_id, name, date, created_by) VALUES (?, ?, ?, ?)",
                    (proj_id, ms["name"], ms["date"], username),
                )
            
            # Add phases
            for ph in proj_data.get("phases", []):
                conn.execute(
                    "INSERT INTO phases (project_id, name, start_date, end_date, created_by) VALUES (?, ?, ?, ?, ?)",
                    (proj_id, ph["name"], ph["start_date"], ph["end_date"], username),
                )
            
            log_action(conn, username, "IMPORT_JSON", f"Imported project '{proj_data['name']}' from JSON")
    
    print(f"✅ Imported {len(data.get('projects', []))} project(s) from {json_path}")


# ─────────────────────────────────────────────────────────────────────────
# Milestone Task Management
# ─────────────────────────────────────────────────────────────────────────

def sync_milestone_tasks_from_kpi(db_path: pathlib.Path | None = None) -> None:
    """Sync tasks from KPI Excel file to database for all milestones."""
    from timeline_tool.kpi_loader import load_milestone_tasks
    
    # Ensure table exists (migration)
    migrate_db(db_path)
    
    kpi_tasks = load_milestone_tasks()
    
    with _connect(db_path) as conn:
        # Get all milestones
        milestones = conn.execute("SELECT id, name FROM milestones").fetchall()
        
        for milestone in milestones:
            ms_id = milestone["id"]
            ms_name = milestone["name"]
            
            # Get tasks for this milestone from Excel
            excel_tasks = kpi_tasks.get(ms_name, [])
            
            # Get existing tasks from database
            existing_tasks = conn.execute(
                "SELECT task_name FROM milestone_tasks WHERE milestone_id = ?",
                (ms_id,)
            ).fetchall()
            existing_task_names = {row["task_name"] for row in existing_tasks}
            
            # Add new tasks from Excel that don't exist in database
            for task_name in excel_tasks:
                if task_name not in existing_task_names:
                    conn.execute(
                        "INSERT INTO milestone_tasks (milestone_id, task_name, status) VALUES (?, ?, ?)",
                        (ms_id, task_name, "Yet to Start")
                    )
        
        conn.commit()


def get_milestone_tasks_with_status(milestone_id: int, db_path: pathlib.Path | None = None) -> list[dict]:
    """Get all tasks for a milestone with their status and attachments."""
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT id, task_name, status, attachment_path FROM milestone_tasks WHERE milestone_id = ? ORDER BY id",
            (milestone_id,)
        ).fetchall()
    return [dict(row) for row in rows]


def update_task_status(task_id: int, status: str, username: str = "system",
                       db_path: pathlib.Path | None = None) -> None:
    """Update the status of a milestone task."""
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE milestone_tasks SET status = ?, updated_by = ?, updated_at = datetime('now') WHERE id = ?",
            (status, username, task_id)
        )
        log_action(conn, username, "UPDATE_TASK_STATUS", f"Updated task {task_id} status to {status}")


def update_task_attachment(task_id: int, attachment_path: str, username: str = "system",
                           db_path: pathlib.Path | None = None) -> None:
    """Update the attachment path for a milestone task."""
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE milestone_tasks SET attachment_path = ?, updated_by = ?, updated_at = datetime('now') WHERE id = ?",
            (attachment_path, username, task_id)
        )
        log_action(conn, username, "UPDATE_TASK_ATTACHMENT", f"Updated task {task_id} attachment")

def remove_task_attachment(task_id: int, username: str = "system",
                           db_path: pathlib.Path | None = None) -> None:
    """Remove the attachment from a milestone task."""
    with _connect(db_path) as conn:
        # Get the current attachment path to delete the file
        row = conn.execute(
            "SELECT attachment_path FROM milestone_tasks WHERE id = ?",
            (task_id,)
        ).fetchone()
        
        if row and row["attachment_path"]:
            # Delete the physical file
            import pathlib
            file_path = pathlib.Path(row["attachment_path"])
            if file_path.exists():
                try:
                    file_path.unlink()
                except Exception as e:
                    print(f"Warning: Could not delete file {file_path}: {e}")
        
        # Clear the attachment path in database
        conn.execute(
            "UPDATE milestone_tasks SET attachment_path = NULL, updated_by = ?, updated_at = datetime('now') WHERE id = ?",
            (username, task_id)
        )
        log_action(conn, username, "REMOVE_TASK_ATTACHMENT", f"Removed attachment from task {task_id}")

def get_qctp_line_items(project_id: int, phase: str, category: str, 
                        db_path: pathlib.Path | None = None) -> list[dict]:
    """Get QCTP line items for a specific project/phase/category."""
    with _connect(db_path) as conn:
        rows = conn.execute(
            """SELECT id, line_number, description, status, remarks, attachment_path
               FROM qctp_line_items 
               WHERE project_id = ? AND phase = ? AND category = ?
               ORDER BY line_number""",
            (project_id, phase, category)
        ).fetchall()
        return [dict(row) for row in rows]


def save_qctp_line_item(project_id: int, phase: str, category: str, 
                        line_number: int, description: str, status: str,
                        remarks: str, attachment_path: str, username: str,
                        db_path: pathlib.Path | None = None) -> int:
    """Save or update a QCTP line item."""
    with _connect(db_path) as conn:
        # Check if exists
        existing = conn.execute(
            """SELECT id FROM qctp_line_items 
               WHERE project_id = ? AND phase = ? AND category = ? AND line_number = ?""",
            (project_id, phase, category, line_number)
        ).fetchone()
        
        if existing:
            conn.execute(
                """UPDATE qctp_line_items 
                   SET description = ?, status = ?, remarks = ?, attachment_path = ?,
                       updated_by = ?, updated_at = datetime('now')
                   WHERE id = ?""",
                (description, status, remarks, attachment_path, username, existing["id"])
            )
            return existing["id"]
        else:
            cursor = conn.execute(
                """INSERT INTO qctp_line_items 
                   (project_id, phase, category, line_number, description, status, 
                    remarks, attachment_path, created_by, updated_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (project_id, phase, category, line_number, description, status,
                 remarks, attachment_path, username, username)
            )
            return cursor.lastrowid

def save_qctp_line_item(project_id: int, phase: str, category: str, 
                        line_number: int, description: str, status: str,
                        remarks: str, attachment_path: str, username: str,
                        db_path: pathlib.Path | None = None) -> int:
    """Save or update a QCTP line item."""
    with _connect(db_path) as conn:
        # Check if exists
        existing = conn.execute(
            """SELECT id FROM qctp_line_items 
               WHERE project_id = ? AND phase = ? AND category = ? AND line_number = ?""",
            (project_id, phase, category, line_number)
        ).fetchone()
        
        if existing:
            conn.execute(
                """UPDATE qctp_line_items 
                   SET description = ?, status = ?, remarks = ?, attachment_path = ?,
                       updated_by = ?, updated_at = datetime('now')
                   WHERE id = ?""",
                (description, status, remarks, attachment_path, username, existing["id"])
            )
            return existing["id"]
        else:
            cursor = conn.execute(
                """INSERT INTO qctp_line_items 
                   (project_id, phase, category, line_number, description, status, 
                    remarks, attachment_path, created_by, updated_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (project_id, phase, category, line_number, description, status,
                 remarks, attachment_path, username, username)
            )
            return cursor.lastrowid

def init_qctp_line_items_from_template(project_id: int, username: str, 
                                        db_path: pathlib.Path | None = None):
    """
    Initialize QCTP line items for a project using the template from Excel.
    Only creates items that don't already exist.
    """
    from timeline_tool.qctp_template import get_qctp_template

    # Load template
    template = load_qctp_template_from_excel()
    
    phases = ["pre_program", "detailed_design", "industrialization"]
    categories = ["quality", "cost", "time", "performance"]
    
    with _connect(db_path) as conn:
        for phase in phases:
            for category in categories:
                for line_num in range(1, 5):
                    # Check if already exists
                    existing = conn.execute(
                        """SELECT id FROM qctp_line_items 
                           WHERE project_id = ? AND phase = ? AND category = ? AND line_number = ?""",
                        (project_id, phase, category, line_num)
                    ).fetchone()
                    
                    if not existing:
                        # Get description from template
                        description = get_qctp_line_item_description(phase, category, line_num, template)
                        
                        conn.execute(
                            """INSERT INTO qctp_line_items 
                               (project_id, phase, category, line_number, description, status, 
                                remarks, attachment_path, updated_by)
                               VALUES (?, ?, ?, ?, ?, 'Green', '', '', ?)""",
                            (project_id, phase, category, line_num, description, username)
                        )
        
        log_action(conn, username, "INIT_QCTP", 
                   f"Initialized QCTP template for project ID {project_id}")
    
    print(f"✅ Initialized QCTP line items for project {project_id}")

def save_qctp_line_item(project_id: int, phase: str, category: str, line_number: int,
                        description: str, status: str, remarks: str, 
                        attachment_path: str, username: str,
                        db_path: pathlib.Path | None = None) -> int:
    """Save or update a QCTP line item."""
    with _connect(db_path) as conn:
        existing = conn.execute(
            """SELECT id FROM qctp_line_items 
               WHERE project_id = ? AND phase = ? AND category = ? AND line_number = ?""",
            (project_id, phase, category, line_number)
        ).fetchone()
        
        if existing:
            conn.execute(
                """UPDATE qctp_line_items 
                   SET description = ?, status = ?, remarks = ?, attachment_path = ?,
                       updated_by = ?, updated_at = datetime('now')
                   WHERE id = ?""",
                (description, status, remarks, attachment_path, username, existing["id"])
            )
            return existing["id"]
        else:
            cursor = conn.execute(
                """INSERT INTO qctp_line_items 
                   (project_id, phase, category, line_number, description, status, 
                    remarks, attachment_path, updated_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (project_id, phase, category, line_number, description, status, 
                 remarks, attachment_path, username)
            )
            return cursor.lastrowid


def init_qctp_line_items(project_id: int, phase: str, category: str,
                         username: str, db_path: pathlib.Path | None = None):
    """Initialize 4 empty line items for a project/phase/category if they don't exist."""
    with _connect(db_path) as conn:
        for line_num in range(1, 5):
            existing = conn.execute(
                """SELECT id FROM qctp_line_items 
                   WHERE project_id = ? AND phase = ? AND category = ? AND line_number = ?""",
                (project_id, phase, category, line_num)
            ).fetchone()
            
            if not existing:
                conn.execute(
                    """INSERT INTO qctp_line_items 
                       (project_id, phase, category, line_number, description, status, 
                        remarks, attachment_path, updated_by)
                       VALUES (?, ?, ?, ?, '', 'Green', '', '', ?)""",
                    (project_id, phase, category, line_num, username)
                )

def get_qctp_week_notes(project_id: int, year: int, week_number: int, 
                        db_path: pathlib.Path | None = None) -> dict:
    """Get QCTP week notes for a project."""
    with _connect(db_path) as conn:
        # Ensure table exists
        conn.execute("""
            CREATE TABLE IF NOT EXISTS qctp_week_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                year INTEGER NOT NULL,
                week_number INTEGER NOT NULL,
                highlights TEXT DEFAULT '',
                red_points TEXT DEFAULT '',
                escalation TEXT DEFAULT '',
                updated_by TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(project_id, year, week_number),
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            )
        """)
        
        row = conn.execute("""
            SELECT highlights, red_points, escalation
            FROM qctp_week_notes
            WHERE project_id = ? AND year = ? AND week_number = ?
        """, (project_id, year, week_number)).fetchone()
        
        if row:
            return {
                "highlights": row["highlights"] or "",
                "red_points": row["red_points"] or "",
                "escalation": row["escalation"] or "",
            }
        return {"highlights": "", "red_points": "", "escalation": ""}


def save_qctp_week_notes(project_id: int, year: int, week_number: int,
                         highlights: str, red_points: str, escalation: str,
                         username: str, db_path: pathlib.Path | None = None):
    """Save QCTP week notes for a project."""
    with _connect(db_path) as conn:
        # Ensure table exists
        conn.execute("""
            CREATE TABLE IF NOT EXISTS qctp_week_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                year INTEGER NOT NULL,
                week_number INTEGER NOT NULL,
                highlights TEXT DEFAULT '',
                red_points TEXT DEFAULT '',
                escalation TEXT DEFAULT '',
                updated_by TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(project_id, year, week_number),
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            )
        """)
        
        conn.execute("""
            INSERT INTO qctp_week_notes (project_id, year, week_number, highlights, red_points, escalation, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(project_id, year, week_number) DO UPDATE SET
                highlights = excluded.highlights,
                red_points = excluded.red_points,
                escalation = excluded.escalation,
                updated_by = excluded.updated_by,
                updated_at = CURRENT_TIMESTAMP
        """, (project_id, year, week_number, highlights, red_points, escalation, username))
        
        conn.commit()

def get_qctp_notes(project_id: int, phase: str, week_number: int, year: int, 
                   db_path: pathlib.Path | None = None) -> dict:
    """Get QCTP notes for a specific project, phase, and week."""
    with _connect(db_path) as conn:
        # Create table if not exists
        conn.execute("""
            CREATE TABLE IF NOT EXISTS qctp_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                phase TEXT NOT NULL,
                week_number INTEGER NOT NULL,
                year INTEGER NOT NULL,
                highlights TEXT DEFAULT '',
                red_points TEXT DEFAULT '',
                escalation TEXT DEFAULT '',
                updated_by TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(project_id, phase, week_number, year)
            )
        """)
        
        row = conn.execute("""
            SELECT highlights, red_points, escalation 
            FROM qctp_notes 
            WHERE project_id = ? AND phase = ? AND week_number = ? AND year = ?
        """, (project_id, phase, week_number, year)).fetchone()
        
        if row:
            return {
                "highlights": row["highlights"] or "",
                "red_points": row["red_points"] or "",
                "escalation": row["escalation"] or "",
            }
        return {}


def save_qctp_notes(project_id: int, phase: str, week_number: int, year: int,
                    highlights: str, red_points: str, escalation: str,
                    username: str, db_path: pathlib.Path | None = None):
    """Save QCTP notes for a specific project, phase, and week."""
    with _connect(db_path) as conn:
        # Create table if not exists
        conn.execute("""
            CREATE TABLE IF NOT EXISTS qctp_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                phase TEXT NOT NULL,
                week_number INTEGER NOT NULL,
                year INTEGER NOT NULL,
                highlights TEXT DEFAULT '',
                red_points TEXT DEFAULT '',
                escalation TEXT DEFAULT '',
                updated_by TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(project_id, phase, week_number, year)
            )
        """)
        
        conn.execute("""
            INSERT INTO qctp_notes (project_id, phase, week_number, year, highlights, red_points, escalation, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(project_id, phase, week_number, year) DO UPDATE SET
                highlights = excluded.highlights,
                red_points = excluded.red_points,
                escalation = excluded.escalation,
                updated_by = excluded.updated_by,
                updated_at = CURRENT_TIMESTAMP
        """, (project_id, phase, week_number, year, highlights, red_points, escalation, username))
        conn.commit()

# ═════════════════════════════════════════════════════════════════════════
# ACTIVITIES
# ═════════════════════════════════════════════════════════════════════════

def add_activity(project_id: int, week_number: int, year: int,
                 activity_name: str, start_date: str, end_date: str,
                 time_taken: str = "", members: str = "",
                 hard_points: str = "", status: str = "WIP",
                 attachment_path: str = "",
                 username: str = "system",
                 db_path: pathlib.Path | None = None) -> int:
    """Add a new activity entry."""
    with _connect(db_path) as conn:
        cursor = conn.execute(
            """INSERT INTO activities
               (project_id, week_number, year, activity_name, start_date, end_date,
                time_taken, members, hard_points, status, attachment_path,
                created_by, updated_by)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (project_id, week_number, year, activity_name, start_date, end_date,
             time_taken, members, hard_points, status, attachment_path,
             username, username)
        )
        activity_id = cursor.lastrowid
        log_action(conn, username, "ADD_ACTIVITY",
                   f"Added activity '{activity_name}' for project {project_id} (W{week_number}/{year})")
    return activity_id


def update_activity(activity_id: int, activity_name: str, start_date: str,
                    end_date: str, time_taken: str = "", members: str = "",
                    hard_points: str = "", status: str = "WIP",
                    attachment_path: str = "",
                    username: str = "system",
                    db_path: pathlib.Path | None = None) -> None:
    """Update an existing activity."""
    with _connect(db_path) as conn:
        conn.execute(
            """UPDATE activities
               SET activity_name = ?, start_date = ?, end_date = ?,
                   time_taken = ?, members = ?, hard_points = ?, status = ?,
                   attachment_path = ?, updated_by = ?, updated_at = datetime('now')
               WHERE id = ?""",
            (activity_name, start_date, end_date, time_taken, members,
             hard_points, status, attachment_path, username, activity_id)
        )
        log_action(conn, username, "UPDATE_ACTIVITY",
                   f"Updated activity ID {activity_id}")


def delete_activity(activity_id: int, username: str = "system",
                    db_path: pathlib.Path | None = None) -> None:
    """Delete an activity."""
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM activities WHERE id = ?", (activity_id,))
        log_action(conn, username, "DELETE_ACTIVITY",
                   f"Deleted activity ID {activity_id}")


def get_activities(project_id: int, week_number: int=None, year: int=None,
                   db_path: pathlib.Path=None) -> list[dict]:
    """Get activities with project name included."""
    with _connect(db_path) as conn:
        if week_number is not None and year is not None:
            rows = conn.execute(
                """
                SELECT a.*, p.name AS project_name
                FROM activities a
                JOIN projects p ON a.project_id = p.id
                WHERE a.project_id = ? AND a.week_number = ? AND a.year = ?
                ORDER BY a.start_date
                """,
                (project_id, week_number, year)
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT a.*, p.name AS project_name
                FROM activities a
                JOIN projects p ON a.project_id = p.id
                WHERE a.project_id = ?
                ORDER BY a.year DESC, a.week_number DESC, a.start_date
                """,
                (project_id,)
            ).fetchall()

        return [dict(r) for r in rows]