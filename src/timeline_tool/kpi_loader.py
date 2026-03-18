"""
Load milestone KPIs/tasks from Excel file.
"""

from __future__ import annotations

import pathlib
import os
from typing import Dict, List

try:
    import openpyxl
except ImportError:
    openpyxl = None


# Dynamic path based on current user
def get_kpi_path() -> pathlib.Path:
    current_user = os.environ.get('USERNAME', 'T0276HS')
    return pathlib.Path(
        f"C:/Users/{current_user}/Stellantis/AI ML - VEHE STRUCTURE - Documents/Shared/project_timelines/milestone_kpi.xlsx"
    )


def load_milestone_tasks(kpi_path: pathlib.Path | None = None) -> Dict[str, List[str]]:
    """
    Load milestone tasks from Excel file.
    Returns a dictionary mapping milestone names to list of tasks.
    
    Expected Excel format:
    | Milestone | Task                     |
    |-----------|--------------------------|
    | IM        | Project Input/Scope      |
    | IM        | Another task for IM      |
    | Other     | Some other task          |
    """
    if openpyxl is None:
        print("⚠️  Cannot load KPI file: openpyxl not installed. Run: pip install openpyxl")
        return {}
    
    path = kpi_path or get_kpi_path()
    
    if not path.exists():
        print(f"⚠️  KPI file not found: {path}")
        return {}
    
    try:
        workbook = openpyxl.load_workbook(path, data_only=True)
        sheet = workbook.active
        
        milestone_tasks: Dict[str, List[str]] = {}
        
        # Skip header row (row 1)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:  # Both milestone and task must exist
                milestone_name = str(row[0]).strip()
                task = str(row[1]).strip()
                
                if milestone_name not in milestone_tasks:
                    milestone_tasks[milestone_name] = []
                
                milestone_tasks[milestone_name].append(task)
        
        print(f"✅ Loaded tasks for {len(milestone_tasks)} milestones from {path.name}")
        return milestone_tasks
        
    except Exception as e:
        print(f"❌ Error loading KPI file: {e}")
        return {}