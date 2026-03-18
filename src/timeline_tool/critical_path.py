"""
Critical Path Analysis for Project Management Tool.
Calculates and visualizes the critical path through project dependencies.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from collections import defaultdict
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from timeline_tool.models import Project, Milestone


@dataclass
class CriticalPathNode:
    """A node in the critical path network."""
    name: str
    duration: int  # in days
    earliest_start: int = 0
    earliest_finish: int = 0
    latest_start: int = 0
    latest_finish: int = 0
    slack: int = 0
    is_critical: bool = False
    predecessors: list[str] = field(default_factory=list)
    successors: list[str] = field(default_factory=list)


@dataclass
class CriticalPathResult:
    """Result of critical path analysis."""
    critical_path: list[str]  # Names of nodes on critical path
    total_duration: int  # Total project duration in days
    nodes: dict[str, CriticalPathNode]  # All nodes
    project_start: datetime.date
    project_end: datetime.date


def calculate_critical_path(
    projects: list,
    use_milestones: bool = True,
) -> CriticalPathResult:
    """
    Calculate the critical path for a set of projects.
    
    The critical path is the longest sequence of dependent activities
    that determines the minimum project duration.
    
    Args:
        projects: List of Project objects with dependencies (depends_on field)
        use_milestones: If True, include milestones as nodes
    
    Returns:
        CriticalPathResult with critical path information
    """
    if not projects:
        return CriticalPathResult(
            critical_path=[],
            total_duration=0,
            nodes={},
            project_start=datetime.date.today(),
            project_end=datetime.date.today()
        )
    
    # Build network nodes from projects
    nodes: dict[str, CriticalPathNode] = {}
    
    # Create a node for each project
    for project in projects:
        node = CriticalPathNode(
            name=project.name,
            duration=project.duration_days(),
            predecessors=list(project.depends_on) if project.depends_on else [],
        )
        nodes[project.name] = node
    
    # Add milestones as sub-nodes if requested
    if use_milestones:
        for project in projects:
            prev_node = project.name
            for i, ms in enumerate(sorted(project.milestones, key=lambda m: m.date)):
                ms_node_name = f"{project.name}::{ms.name}"
                
                # Calculate milestone duration from previous milestone or project start
                if i == 0:
                    ms_duration = (ms.date - project.start_date).days
                else:
                    prev_ms = sorted(project.milestones, key=lambda m: m.date)[i-1]
                    ms_duration = (ms.date - prev_ms.date).days
                
                ms_node = CriticalPathNode(
                    name=ms_node_name,
                    duration=max(0, ms_duration),
                    predecessors=[prev_node],
                )
                nodes[ms_node_name] = ms_node
                prev_node = ms_node_name
    
    # Build successor relationships
    for name, node in nodes.items():
        for pred_name in node.predecessors:
            if pred_name in nodes:
                nodes[pred_name].successors.append(name)
    
    # Forward pass: calculate earliest start and finish
    # Process nodes in topological order
    processed = set()
    
    def forward_pass(node_name: str):
        if node_name in processed:
            return
        
        node = nodes.get(node_name)
        if not node:
            return
        
        # Process all predecessors first
        for pred_name in node.predecessors:
            if pred_name not in processed:
                forward_pass(pred_name)
        
        # Calculate earliest start (max of all predecessor finish times)
        if node.predecessors:
            pred_finishes = [
                nodes[p].earliest_finish 
                for p in node.predecessors 
                if p in nodes
            ]
            node.earliest_start = max(pred_finishes) if pred_finishes else 0
        else:
            node.earliest_start = 0
        
        node.earliest_finish = node.earliest_start + node.duration
        processed.add(node_name)
    
    # Run forward pass on all nodes
    for name in nodes:
        forward_pass(name)
    
    # Find total project duration
    total_duration = max(n.earliest_finish for n in nodes.values()) if nodes else 0
    
    # Backward pass: calculate latest start and finish
    # Start from nodes with no successors
    
    def backward_pass(node_name: str, visited: set):
        if node_name in visited:
            return
        
        node = nodes.get(node_name)
        if not node:
            return
        
        # Process all successors first
        for succ_name in node.successors:
            if succ_name not in visited:
                backward_pass(succ_name, visited)
        
        # Calculate latest finish
        if node.successors:
            succ_starts = [
                nodes[s].latest_start 
                for s in node.successors 
                if s in nodes
            ]
            node.latest_finish = min(succ_starts) if succ_starts else total_duration
        else:
            node.latest_finish = total_duration
        
        node.latest_start = node.latest_finish - node.duration
        visited.add(node_name)
    
    # Run backward pass
    visited = set()
    for name in nodes:
        backward_pass(name, visited)
    
    # Calculate slack and identify critical path
    critical_path = []
    for name, node in nodes.items():
        node.slack = node.latest_start - node.earliest_start
        node.is_critical = (node.slack == 0)
        if node.is_critical:
            critical_path.append(name)
    
    # Sort critical path by earliest start
    critical_path.sort(key=lambda n: nodes[n].earliest_start)
    
    # Calculate actual dates
    earliest_date = min(p.start_date for p in projects)
    latest_date = earliest_date + datetime.timedelta(days=total_duration)
    
    return CriticalPathResult(
        critical_path=critical_path,
        total_duration=total_duration,
        nodes=nodes,
        project_start=earliest_date,
        project_end=latest_date,
    )


def get_critical_path_summary(result: CriticalPathResult) -> str:
    """
    Generate a text summary of the critical path analysis.
    """
    lines = [
        "=" * 60,
        "CRITICAL PATH ANALYSIS",
        "=" * 60,
        f"Project Duration: {result.total_duration} days",
        f"Start Date: {result.project_start}",
        f"End Date: {result.project_end}",
        "",
        "CRITICAL PATH:",
        "-" * 40,
    ]
    
    for i, node_name in enumerate(result.critical_path, 1):
        node = result.nodes[node_name]
        lines.append(f"  {i}. {node_name}")
        lines.append(f"     Duration: {node.duration} days")
        lines.append(f"     ES: Day {node.earliest_start} → EF: Day {node.earliest_finish}")
    
    lines.extend([
        "",
        "ALL ACTIVITIES:",
        "-" * 40,
    ])
    
    # Sort by earliest start
    sorted_nodes = sorted(result.nodes.items(), key=lambda x: x[1].earliest_start)
    for name, node in sorted_nodes:
        critical_mark = " ⭐" if node.is_critical else ""
        lines.append(f"  {name}{critical_mark}")
        lines.append(f"    Duration: {node.duration}d | Slack: {node.slack}d")
        lines.append(f"    ES: {node.earliest_start} | EF: {node.earliest_finish} | "
                    f"LS: {node.latest_start} | LF: {node.latest_finish}")
    
    return "\n".join(lines)


def highlight_critical_path_on_chart(
    ax,
    projects: list,
    critical_path_result: CriticalPathResult,
    y_positions: dict[str, float],
    highlight_color: str = "#E74C3C",
    highlight_alpha: float = 0.3,
) -> None:
    """
    Highlight the critical path on an existing matplotlib Gantt chart.
    
    Args:
        ax: Matplotlib axes object
        projects: List of projects
        critical_path_result: Result from calculate_critical_path()
        y_positions: Dict mapping project names to Y positions
        highlight_color: Color for critical path highlighting
        highlight_alpha: Alpha for highlight overlay
    """
    import matplotlib.patches as mpatches
    import matplotlib.dates as mdates
    
    for project in projects:
        if project.name not in critical_path_result.critical_path:
            continue
        
        if project.name not in y_positions:
            continue
        
        y = y_positions[project.name]
        node = critical_path_result.nodes[project.name]
        
        # Add highlight rectangle
        start_num = mdates.date2num(project.start_date)
        end_num = mdates.date2num(project.end_date)
        width = end_num - start_num
        
        rect = mpatches.Rectangle(
            (start_num, y - 0.2),
            width,
            0.4,
            linewidth=3,
            edgecolor=highlight_color,
            facecolor=highlight_color,
            alpha=highlight_alpha,
            zorder=2,
        )
        ax.add_patch(rect)
        
        # Add "CRITICAL" label
        ax.text(
            start_num + width / 2,
            y + 0.25,
            "CRITICAL",
            fontsize=7,
            ha="center",
            va="bottom",
            color=highlight_color,
            fontweight="bold",
            zorder=10,
        )