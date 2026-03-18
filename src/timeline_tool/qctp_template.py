"""
QCTP Template loader - loads line items from Excel file.
"""

import pathlib
from typing import Dict, List

# Default QCTP template path - use raw string to avoid unicode escape issues
QCTP_TEMPLATE_PATH = pathlib.Path(
    r"C:\Users\T0276HS\Stellantis\AI ML - VEHE STRUCTURE - Documents\Shared\project_timelines\qctp.xlsx"
)

# Hardcoded template data based on the Excel structure (fallback if file not available)
QCTP_TEMPLATE = {
    "pre_program": {
        "quality": [
            "Quality Must Have + LL NCBS Quality in use (QIU)",
            "PORO - Technical Robustness+ TPC + Capex (G/R)",
            "R&D",
            "FTE plan ED&D",
        ],
        "cost": [
            "ECONOMICS",
            "Validation plan",
            "budget",
            "Sourcing Plan",
        ],
        "time": [
            "Milestone Readiness",
            "KBE sections + archetype",
            "Architecture convergence",
            "Perfo (spider target + scenary)",
        ],
        "performance": [
            "Weight Target",
            "Benchmark Reduction proposals",
            "",
            "",
        ],
    },
    "detailed_design": {
        "quality": [
            "Design Quality DQM Design integrity / Perceived quality / G&F book",
            "APQP Gate 1&2",
            "Sourcing Status",
            "MANUFACTURING Design to Manufacturing DTM",
        ],
        "cost": [
            "Economics (Gap vs PORO, design to cost, BIW: WSE, n of parts, Coeficient material, buffles)",
            "",
            "",
            "",
        ],
        "time": [
            "CAD readiness (Sync 3 / Sync5)",
            "SHRM",
            "Part Readiness (OT / OTOP)",
            "",
        ],
        "performance": [
            "Weight",
            "Style Convergence (STPI + B class + A class)",
            "Packaging (interfaces+ Harness)",
            "Performance convergence (status vs target, planning)",
        ],
    },
    "industrialization": {
        "quality": [
            "DVP DVPa / DPVf / DVPw / conformity of parts / Redesign / BIW Geo",
            "Interim containment action (ICA) Courbe forecast to 0",
            "Manufacturing launch KPI wPi",
            "",
        ],
        "cost": [
            "Economics status of TPC and Investmets",
            "ECRs / CN ODM Status (qty + timing)",
            "",
            "",
        ],
        "time": [
            "Timing and Readiness (maplex / OT / OTOP / PAPP A)",
            "Validation validation plan (subsystem + component)",
            "Performance status, result vs forecast",
            "Homologation status",
        ],
        "performance": [
            "Weight",
            "",
            "",
            "",
        ],
    },
}


def load_qctp_template_from_excel(filepath: pathlib.Path = None) -> Dict[str, Dict[str, List[str]]]:
    """
    Load QCTP template from Excel file.
    
    Returns a dictionary structured as:
    {
        "pre_program": {
            "quality": ["item1", "item2", "item3", "item4"],
            ...
        },
        ...
    }
    """
    if filepath is None:
        filepath = QCTP_TEMPLATE_PATH
    
    if not filepath.exists():
        print(f"QCTP template file not found: {filepath}")
        print("Using hardcoded template data.")
        return QCTP_TEMPLATE
    
    try:
        import openpyxl
    except ImportError:
        print("openpyxl not available. Using hardcoded template data.")
        return QCTP_TEMPLATE
    
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        
        template = {
            "pre_program": {"quality": [], "cost": [], "time": [], "performance": []},
            "detailed_design": {"quality": [], "cost": [], "time": [], "performance": []},
            "industrialization": {"quality": [], "cost": [], "time": [], "performance": []},
        }
        
        # Column mappings based on the Excel structure (1-indexed)
        # Pre-program: B=2, C=3, D=4, E=5
        # Detailed Design: F=6, G=7, H=8, I=9
        # Industrialization: J=10, K=11, L=12, M=13
        phase_columns = {
            "pre_program": {"quality": 2, "cost": 3, "time": 4, "performance": 5},
            "detailed_design": {"quality": 6, "cost": 7, "time": 8, "performance": 9},
            "industrialization": {"quality": 10, "cost": 11, "time": 12, "performance": 13},
        }
        
        # Read data rows (rows 5-8)
        for row_idx in range(5, 9):
            for phase, columns in phase_columns.items():
                for category, col_idx in columns.items():
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    template[phase][category].append(str(cell_value).strip() if cell_value else "")
        
        wb.close()
        
        # Ensure each category has exactly 4 items
        for phase in template:
            for category in template[phase]:
                while len(template[phase][category]) < 4:
                    template[phase][category].append("")
                template[phase][category] = template[phase][category][:4]
        
        print(f"Loaded QCTP template from: {filepath}")
        return template
        
    except Exception as e:
        print(f"Error loading QCTP template from Excel: {e}")
        print("Using hardcoded template data.")
        return QCTP_TEMPLATE


def get_qctp_template() -> Dict[str, Dict[str, List[str]]]:
    """Get the QCTP template, loading from Excel if available."""
    return load_qctp_template_from_excel()


def get_qctp_line_item_description(phase: str, category: str, line_number: int,
                                    template: Dict = None) -> str:
    """Get the default description for a QCTP line item."""
    if template is None:
        template = QCTP_TEMPLATE
    
    try:
        items = template.get(phase, {}).get(category, [])
        if 0 <= line_number - 1 < len(items):
            return items[line_number - 1]
    except (KeyError, IndexError):
        pass
    
    return ""