"""
First-time setup: initialise the database and create the admin user.
Run this ONCE before using the tool.

Usage:
    python run_admin.py
"""

import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from timeline_tool.database import init_db, import_from_json, DEFAULT_DB_PATH
from timeline_tool.auth import create_user

print("=" * 60)
print("  📊 Project Timeline Tool — First-Time Setup")
print("=" * 60)
print(f"\n  Database location: {DEFAULT_DB_PATH}\n")

# Step 1: Initialise DB
init_db()

# Step 2: Create admin user
print("\n--- Create Admin Account ---")
username = input("  Admin username: ").strip()
password = input("  Admin password: ").strip()
full_name = input("  Full name (optional): ").strip()

create_user(username, password, role="admin", full_name=full_name)

# Step 3: Optionally import existing JSON data
print()
import_data = input("  Import data from sample_projects.json? (y/n): ").strip().lower()
if import_data == "y":
    json_path = os.path.join(PROJECT_ROOT, "data", "sample_projects.json")
    if os.path.exists(json_path):
        import_from_json(json_path, username=username)
    else:
        print(f"  ⚠️  File not found: {json_path}")

print("\n" + "=" * 60)
print("  ✅ Setup complete! Run 'python run.py' to launch the tool.")
print("=" * 60)