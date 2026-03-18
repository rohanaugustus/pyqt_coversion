"""
Default configuration for the timeline renderer.
"""

from typing import Final

# ── Colour palette (when project has no explicit colour) ─────────────
DEFAULT_COLORS: Final[list[str]] = [
    "#4C72B0", "#55A868", "#C44E52", "#8172B2",
    "#CCB974", "#64B5CD", "#D48E2A", "#8C8C8C",
]

# ── Default project color ─────────────────────────────────────────────
DEFAULT_PROJECT_COLOR: Final[str] = "#243782"

# ── Status‑based colours (Enhancement #2) ───────────────────────────
STATUS_COLORS: Final[dict[str, str]] = {
    "on-track": "#27AE60",   # green
    "at-risk":  "#F39C12",   # amber
    "overdue":  "#E74C3C",   # red
}
STATUS_BORDER_WIDTH: Final[float] = 2.5

# ── Today‑line styling ──────────────────────────────────────────────
TODAY_LINE_COLOR: Final[str] = "#E74C3C"
TODAY_LINE_WIDTH: Final[float] = 2.0
TODAY_LINE_STYLE: Final[str] = "--"

# ── Bar styling ─────────────────────────────────────────────────────
BAR_HEIGHT: Final[float] = 0.3
PHASE_HEIGHT: Final[float] = 0.14   # thinner nested bars for phases

# ── Milestone styling ───────────────────────────────────────────────
MILESTONE_MARKER: Final[str] = "o"
MILESTONE_SIZE: Final[int] = 70
MILESTONE_EDGE_COLOR: Final[str] = "black"

# ── Dependency arrow styling ────────────────────────────────────────
DEP_ARROW_COLOR: Final[str] = "#555555"
DEP_ARROW_WIDTH: Final[float] = 1.5
DEP_ARROW_STYLE: Final[str] = "->"

# ── Figure defaults ─────────────────────────────────────────────────
FIGURE_WIDTH: Final[int] = 20
FIGURE_HEIGHT_PER_PROJECT: Final[float] = 0.9
FIGURE_MIN_HEIGHT: Final[float] = 5.0
DPI: Final[int] = 150
TITLE_FONT_SIZE: Final[int] = 16
LABEL_FONT_SIZE: Final[int] = 10
DATE_FORMAT: Final[str] = "%Y-%m-%d"

# ── Theme (Enhancement #5) ──────────────────────────────────────────
THEMES: Final[dict[str, dict]] = {
    "light": {
        "bg_color":       "#FFFFFF",
        "text_color":     "#333333",
        "grid_color":     "#CCCCCC",
        "axis_bg":        "#FAFAFA",
        "date_label_color": "#555555",
        "tooltip_bg":     "#333333",
        "tooltip_text":   "white",
    },
    "dark": {
        "bg_color":       "#1E1E1E",
        "text_color":     "#E0E0E0",
        "grid_color":     "#444444",
        "axis_bg":        "#2A2A2A",
        "date_label_color": "#AAAAAA",
        "tooltip_bg":     "#E0E0E0",
        "tooltip_text":   "#1E1E1E",
    },
}