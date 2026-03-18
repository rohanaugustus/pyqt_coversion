"""
User authentication and role-based access control.

Roles
─────
  admin  → Can manage users, edit all data, view dashboard
  editor → Can edit project data, milestones, phases; view dashboard
  viewer → Can only view the dashboard (read-only)
"""

from __future__ import annotations

import sqlite3
import pathlib

import bcrypt

from timeline_tool.database import _connect, DEFAULT_DB_PATH, log_action


# ─────────────────────────────────────────────────────────────────────────
# Role definitions
# ─────────────────────────────────────────────────────────────────────────

ROLES = {
    "admin":  {"can_edit": True, "can_manage_users": True,  "can_view": True},
    "editor": {"can_edit": True, "can_manage_users": False, "can_view": True},
    "viewer": {"can_edit": False, "can_manage_users": False, "can_view": True},
}


def _hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def _check_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode("utf-8"), hashed.encode("utf-8"))


# ─────────────────────────────────────────────────────────────────────────
# User management
# ─────────────────────────────────────────────────────────────────────────

def create_user(
    username: str, password: str, role: str = "viewer",
    full_name: str = "", db_path: pathlib.Path | None = None,
) -> None:
    if role not in ROLES:
        raise ValueError(f"Invalid role '{role}'. Must be one of: {list(ROLES.keys())}")
    hashed = _hash_password(password)
    with _connect(db_path) as conn:
        try:
            conn.execute(
                "INSERT INTO users (username, password, role, full_name) VALUES (?, ?, ?, ?)",
                (username, hashed, role, full_name),
            )
            log_action(conn, "system", "CREATE_USER", f"Created user '{username}' with role '{role}'")
        except sqlite3.IntegrityError:
            raise ValueError(f"User '{username}' already exists.")
    print(f"✅ User '{username}' created with role '{role}'")


def authenticate(username: str, password: str, db_path: pathlib.Path | None = None) -> dict | None:
    """
    Authenticate a user. Returns a dict with user info if successful, None otherwise.
    """
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT username, password, role, full_name FROM users WHERE username = ?",
            (username,),
        ).fetchone()

    if row is None:
        return None
    if not _check_password(password, row["password"]):
        return None

    return {
        "username": row["username"],
        "role": row["role"],
        "full_name": row["full_name"],
        "permissions": ROLES[row["role"]],
    }


def list_users(db_path: pathlib.Path | None = None) -> list[dict]:
    with _connect(db_path) as conn:
        rows = conn.execute("SELECT username, role, full_name, created_at FROM users ORDER BY username").fetchall()
    return [dict(r) for r in rows]


def update_user_role(username: str, new_role: str, admin_user: str = "system",
                     db_path: pathlib.Path | None = None) -> None:
    if new_role not in ROLES:
        raise ValueError(f"Invalid role '{new_role}'.")
    with _connect(db_path) as conn:
        conn.execute("UPDATE users SET role = ? WHERE username = ?", (new_role, username))
        log_action(conn, admin_user, "UPDATE_ROLE", f"Changed '{username}' role to '{new_role}'")
    print(f"✅ User '{username}' role changed to '{new_role}'")


def delete_user(username: str, admin_user: str = "system", db_path: pathlib.Path | None = None) -> None:
    with _connect(db_path) as conn:
        conn.execute("DELETE FROM users WHERE username = ?", (username,))
        log_action(conn, admin_user, "DELETE_USER", f"Deleted user '{username}'")
    print(f"✅ User '{username}' deleted")


def change_password(username: str, new_password: str, db_path: pathlib.Path | None = None) -> None:
    hashed = _hash_password(new_password)
    with _connect(db_path) as conn:
        conn.execute("UPDATE users SET password = ? WHERE username = ?", (hashed, username))
        log_action(conn, username, "CHANGE_PASSWORD", f"Password changed for '{username}'")
    print(f"✅ Password changed for '{username}'")