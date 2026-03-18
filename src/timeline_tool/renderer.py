"""
Render a Gantt‑style timeline chart with Matplotlib.

Enhancements included
─────────────────────
#1  CSV / Excel input          → handled in loader.py
#2  Colour‑by‑status          → status border on bars
#3  Phase / sub‑task bars      → thin nested bars inside project bars
#4  Interactive HTML export    → --html flag (Plotly)
#5  Dark mode                  → --theme dark
#6  Dependency arrows          → curved arrows between dependent projects
#7  Custom reference lines     → additional vertical lines (e.g. Board Meeting)
"""

from __future__ import annotations

import datetime
import pathlib

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch

from timeline_tool import config as cfg
from timeline_tool.models import Project, ReferenceLine
from timeline_tool.utils import date_range_padded


# ─────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────

def _pick_color(index: int, project: Project) -> str:
    if project.color:
        return project.color
    return cfg.DEFAULT_COLORS[index % len(cfg.DEFAULT_COLORS)]


def _status_edge_color(project: Project) -> str:
    return cfg.STATUS_COLORS.get(project.status, cfg.STATUS_COLORS["on-track"])


def _apply_theme(fig, ax, theme_name: str) -> dict:
    theme = cfg.THEMES.get(theme_name, cfg.THEMES["light"])
    fig.patch.set_facecolor(theme["bg_color"])
    ax.set_facecolor(theme["axis_bg"])
    ax.title.set_color(theme["text_color"])
    ax.xaxis.label.set_color(theme["text_color"])
    ax.yaxis.label.set_color(theme["text_color"])
    ax.tick_params(colors=theme["text_color"])
    for spine in ax.spines.values():
        spine.set_color(theme["grid_color"])
    return theme


PHASE_PALETTE = ["#85C1E9", "#82E0AA", "#F8C471", "#D7BDE2", "#F0B27A", "#AED6F1"]


# ─────────────────────────────────────────────────────────────────────────
# Main render function
# ─────────────────────────────────────────────────────────────────────────

def render_timeline(
    projects: list[Project],
    today: datetime.date | None = None,
    title: str = "Project Timelines",
    output_path: str | pathlib.Path | None = None,
    show: bool = True,
    theme: str = "light",
    reference_lines: list[ReferenceLine] | None = None,
    html_path: str | pathlib.Path | None = None,
) -> None:
    today = today or datetime.date.today()
    reference_lines = reference_lines or []
    n = len(projects)

    # ── Increase vertical spacing per project to avoid overlap ───────────
    fig_h = max(cfg.FIGURE_MIN_HEIGHT, cfg.FIGURE_HEIGHT_PER_PROJECT * n + 3)
    fig, ax = plt.subplots(figsize=(cfg.FIGURE_WIDTH, fig_h))

    theme_cfg = _apply_theme(fig, ax, theme)

    x_min, x_max = date_range_padded(projects)
    y_min_orig = 0
    y_max_orig = n + 1

    y_positions: list[float] = []
    y_labels: list[str] = []
    milestone_artists = []
    project_y_map: dict[str, float] = {}

    for idx, project in enumerate(projects):
        y = n - idx
        y_positions.append(y)
        y_labels.append(project.name)
        project_y_map[project.name] = y
        color = _pick_color(idx, project)
        status_color = _status_edge_color(project)

        # ── Full bar (background) ────────────────────────────────────────
        ax.barh(
            y,
            (project.end_date - project.start_date).days,
            left=project.start_date,
            height=cfg.BAR_HEIGHT,
            color=color,
            alpha=0.30,
            edgecolor="none",
            linewidth=0,
        )

        # ── Progress overlay ─────────────────────────────────────────────
        progress = project.progress(today)
        if progress > 0:
            elapsed_days = (project.end_date - project.start_date).days * progress
            ax.barh(
                y,
                elapsed_days,
                left=project.start_date,
                height=cfg.BAR_HEIGHT,
                color=color,
                alpha=0.85,
                edgecolor="none",
                linewidth=0,
            )

        # ── Phase sub‑bars (Enhancement #3) ──────────────────────────────
        for p_idx, phase in enumerate(project.phases):
            phase_color = PHASE_PALETTE[p_idx % len(PHASE_PALETTE)]
            phase_days = (phase.end_date - phase.start_date).days
            ax.barh(
                y,
                phase_days,
                left=phase.start_date,
                height=cfg.PHASE_HEIGHT,
                color=phase_color,
                alpha=0.75,
                edgecolor="#666666",
                linewidth=0.5,
                zorder=3,
            )
            # Phase label — positioned BELOW the bar to avoid milestone overlap
            if phase_days > 45:
                mid = phase.start_date + datetime.timedelta(days=phase_days / 2)
                ax.text(
                    mid,
                    y - cfg.BAR_HEIGHT / 2 - 0.05,
                    phase.name,
                    fontsize=cfg.LABEL_FONT_SIZE - 3,
                    ha="center", va="top",
                    color="#333333", fontweight="bold",
                    zorder=4,
                )

        # ── Date labels ──────────────────────────────────────────────────
        ax.text(
            project.start_date,
            y - cfg.BAR_HEIGHT / 2 - 0.22,
            project.start_date.strftime("%b %d, %Y"),
            fontsize=cfg.LABEL_FONT_SIZE - 2,
            ha="left", va="top",
            color=theme_cfg["date_label_color"],
        )
        ax.text(
            project.end_date,
            y - cfg.BAR_HEIGHT / 2 - 0.22,
            project.end_date.strftime("%b %d, %Y"),
            fontsize=cfg.LABEL_FONT_SIZE - 2,
            ha="right", va="top",
            color=theme_cfg["date_label_color"],
        )

        # ── Status indicator label ───────────────────────────────────────
        status_label = project.status.replace("-", " ").upper()
        ax.text(
            project.start_date - datetime.timedelta(days=5),
            y + cfg.BAR_HEIGHT / 2 + 0.1,
            f"● {status_label}",
            fontsize=cfg.LABEL_FONT_SIZE - 3,
            ha="right", va="bottom",
            color=status_color,
            fontweight="bold",
        )

        # ── Milestones ───────────────────────────────────────────────────
        for ms_idx, ms in enumerate(project.milestones):
            # 4-color milestone logic
            milestone_color = ms.marker_color(today)

            sc = ax.scatter(
                ms.date, y,
                marker=cfg.MILESTONE_MARKER,
                s=cfg.MILESTONE_SIZE,
                color=milestone_color,
                edgecolors=cfg.MILESTONE_EDGE_COLOR,
                linewidths=0.8, zorder=5,
            )
            milestone_artists.append((sc, project.name, ms.name, ms.date))

            # Alternate label position above/below to avoid overlap
            if ms_idx % 2 == 0:
                label_y = y + cfg.BAR_HEIGHT * 0.6
                va = "bottom"
            else:
                label_y = y - cfg.BAR_HEIGHT * 0.6
                va = "top"

            ax.text(
                mdates.date2num(ms.date), label_y, ms.name,
                fontsize=cfg.LABEL_FONT_SIZE - 3,
                ha="center", va=va,
                color=theme_cfg["text_color"],
                fontweight="bold", zorder=7,
            )

    # ── Dependency arrows (Enhancement #6) ───────────────────────────────
    for project in projects:
        if not project.depends_on:
            continue
        target_y = project_y_map[project.name]
        for dep_name in project.depends_on:
            if dep_name not in project_y_map:
                continue
            source_y = project_y_map[dep_name]
            source_proj = next((p for p in projects if p.name == dep_name), None)
            if source_proj is None:
                continue

            arrow = FancyArrowPatch(
                posA=(mdates.date2num(source_proj.end_date), source_y),
                posB=(mdates.date2num(project.start_date), target_y),
                arrowstyle=cfg.DEP_ARROW_STYLE,
                color=cfg.DEP_ARROW_COLOR,
                linewidth=cfg.DEP_ARROW_WIDTH,
                connectionstyle="arc3,rad=0.2",
                zorder=6,
            )
            ax.add_patch(arrow)

    # ── Today line ───────────────────────────────────────────────────────
    ax.axvline(
        today,
        color=cfg.TODAY_LINE_COLOR,
        linewidth=cfg.TODAY_LINE_WIDTH,
        linestyle=cfg.TODAY_LINE_STYLE,
        label=f"Today ({today.strftime('%b %d, %Y')})",
        zorder=4,
    )

    # ── Today date label at top of the vertical line ─────────────────────
    ax.text(
        mdates.date2num(today), y_max_orig - 0.05,
        today.strftime("%b %d, %Y"),
        fontsize=cfg.LABEL_FONT_SIZE - 2,
        fontweight="bold",
        color=cfg.TODAY_LINE_COLOR,
        ha="center", va="top",
        zorder=8,
        bbox=dict(boxstyle="round,pad=0.2", facecolor="white",
                  edgecolor=cfg.TODAY_LINE_COLOR, alpha=0.85),
    )

    # ── Custom reference lines (Enhancement #7) ─────────────────────────
    for ref in reference_lines:
        ax.axvline(
            ref.date,
            color=ref.color,
            linewidth=1.8,
            linestyle=ref.style,
            zorder=4,
        )
        ax.text(
            ref.date, y_max_orig - 0.15,
            f"  {ref.name}\n  {ref.date.strftime('%b %d, %Y')}",
            fontsize=cfg.LABEL_FONT_SIZE - 2,
            color=ref.color,
            fontweight="bold",
            ha="left", va="top",
        )

    # ── Axes formatting ──────────────────────────────────────────────────
    ax.set_xlim(x_min, x_max)
    ax.set_ylim(y_min_orig, y_max_orig)
    ax.set_yticks(y_positions)
    ax.set_yticklabels(y_labels, fontsize=cfg.LABEL_FONT_SIZE)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_minor_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    fig.autofmt_xdate(rotation=45, ha="right")
    ax.grid(axis="x", linestyle=":", linewidth=0.5, alpha=0.6, color=theme_cfg["grid_color"])
    ax.set_axisbelow(True)
    ax.set_title(title, fontsize=cfg.TITLE_FONT_SIZE, fontweight="bold", pad=20)
    ax.set_xlabel("Date", fontsize=cfg.LABEL_FONT_SIZE)

    # ── Legend ────────────────────────────────────────────────────────────
    legend_handles = []
    legend_handles.append(
        plt.Line2D([0], [0], color=cfg.TODAY_LINE_COLOR,
                   linewidth=cfg.TODAY_LINE_WIDTH, linestyle=cfg.TODAY_LINE_STYLE,
                   label=f"Today ({today.strftime('%b %d, %Y')})")
    )
    for ref in reference_lines:
        legend_handles.append(
            plt.Line2D([0], [0], color=ref.color, linewidth=1.8,
                       linestyle=ref.style, label=ref.name)
        )
    for status, scolor in cfg.STATUS_COLORS.items():
        legend_handles.append(
            plt.Line2D([0], [0], marker="s", color="w", markerfacecolor=scolor,
                       markersize=8, label=status.replace("-", " ").title())
        )
    # Milestone color legend
    milestone_legend = [
        ("#27AE60", "Completed"),
        ("#F39C12", "At Risk (≤5 days)"),
        ("#E74C3C", "Overdue"),
        ("#5DADE2", "Upcoming"),
    ]
    for ms_color, ms_label in milestone_legend:
        legend_handles.append(
            plt.Line2D([0], [0], marker=cfg.MILESTONE_MARKER, color="w",
                       markerfacecolor=ms_color, markeredgecolor=cfg.MILESTONE_EDGE_COLOR,
                       markersize=10, label=ms_label)
        )

    ax.legend(
        handles=legend_handles,
        loc="upper right",
        fontsize=cfg.LABEL_FONT_SIZE - 1,
        framealpha=0.9,
        facecolor=theme_cfg["bg_color"],
        labelcolor=theme_cfg["text_color"],
    )

    plt.tight_layout()

    # ── Save static image ────────────────────────────────────────────────
    if output_path:
        out = pathlib.Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out, dpi=cfg.DPI, bbox_inches="tight", facecolor=fig.get_facecolor())
        print(f"✅ Chart saved to {out}")

    # ── Interactive HTML export (Enhancement #4) ─────────────────────────
    if html_path:
        _export_html(projects, today, title, reference_lines, html_path)

    # ── Interactive features (hover + zoom) ──────────────────────────────
    if show:
        tooltip = ax.annotate(
            "", xy=(0, 0), xytext=(15, 15),
            textcoords="offset points",
            fontsize=cfg.LABEL_FONT_SIZE, fontweight="bold",
            color=theme_cfg["tooltip_text"],
            bbox=dict(boxstyle="round,pad=0.5", facecolor=theme_cfg["tooltip_bg"],
                      edgecolor="grey", alpha=0.9),
            arrowprops=dict(arrowstyle="->", color=theme_cfg["tooltip_bg"], lw=1.5),
            zorder=10, visible=False,
        )

        def on_hover(event):
            if event.inaxes != ax:
                if tooltip.get_visible():
                    tooltip.set_visible(False)
                    fig.canvas.draw_idle()
                return
            found = False
            for sc, proj_name, ms_name, ms_date in milestone_artists:
                contains, _ = sc.contains(event)
                if contains:
                    tooltip.xy = (mdates.date2num(ms_date), sc.get_offsets()[0][1])
                    tooltip.set_text(f"{proj_name}\n{ms_name}: {ms_date.strftime('%b %d, %Y')}")
                    tooltip.set_visible(True)
                    found = True
                    break
            if not found and tooltip.get_visible():
                tooltip.set_visible(False)
            fig.canvas.draw_idle()

        ZOOM_FACTOR = 0.15
        x_min_num = mdates.date2num(x_min)
        x_max_num = mdates.date2num(x_max)

        def on_scroll(event):
            if event.inaxes != ax:
                return
            cur_x_min, cur_x_max = ax.get_xlim()
            cur_y_min, cur_y_max = ax.get_ylim()
            cx, cy = event.xdata, event.ydata

            if event.button == "up":
                scale = 1 - ZOOM_FACTOR
            elif event.button == "down":
                scale = 1 + ZOOM_FACTOR
            else:
                return

            new_xw = (cur_x_max - cur_x_min) * scale
            rx = (cx - cur_x_min) / (cur_x_max - cur_x_min) if (cur_x_max - cur_x_min) else 0.5
            new_x_min = cx - rx * new_xw
            new_x_max = cx + (1 - rx) * new_xw
            new_x_min = max(new_x_min, x_min_num)
            new_x_max = min(new_x_max, x_max_num)
            if new_x_max - new_x_min < 7:
                return

            new_yw = (cur_y_max - cur_y_min) * scale
            ry = (cy - cur_y_min) / (cur_y_max - cur_y_min) if (cur_y_max - cur_y_min) else 0.5
            new_y_min = cy - ry * new_yw
            new_y_max = cy + (1 - ry) * new_yw
            new_y_min = max(new_y_min, y_min_orig)
            new_y_max = min(new_y_max, y_max_orig)
            if new_y_max - new_y_min < 1.5:
                return

            ax.set_xlim(new_x_min, new_x_max)
            ax.set_ylim(new_y_min, new_y_max)
            fig.canvas.draw_idle()

        def on_click(event):
            if event.inaxes != ax:
                return
            if event.button == 2:
                ax.set_xlim(x_min_num, x_max_num)
                ax.set_ylim(y_min_orig, y_max_orig)
                fig.canvas.draw_idle()

        fig.canvas.mpl_connect("motion_notify_event", on_hover)
        fig.canvas.mpl_connect("scroll_event", on_scroll)
        fig.canvas.mpl_connect("button_press_event", on_click)

        print("\n🖱️  Controls:")
        print("   Scroll Up    → Zoom in (toward cursor)")
        print("   Scroll Down  → Zoom out (from cursor)")
        print("   Middle Click → Reset to full view")
        print("   Hover ◆      → Show milestone date")

        plt.show()

    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────
# Enhancement #4 — Interactive HTML export via Plotly
# ─────────────────────────────────────────────────────────────────────────

def _export_html(
    projects: list[Project],
    today: datetime.date,
    title: str,
    reference_lines: list[ReferenceLine],
    html_path: str | pathlib.Path,
) -> None:
    try:
        import plotly.graph_objects as go
    except ImportError:
        print("⚠️  Plotly not installed. Run: pip install plotly")
        print("   Skipping HTML export.")
        return

    fig = go.Figure()

    for idx, project in enumerate(projects):
        color = _pick_color(idx, project)
        status_color = _status_edge_color(project)

        start_dt = datetime.datetime.combine(project.start_date, datetime.time())
        duration_days = (project.end_date - project.start_date).days

        fig.add_trace(go.Bar(
            x=[duration_days * 86400000],
            y=[project.name],
            base=[start_dt],
            orientation="h",
            marker=dict(color=color, opacity=0.5,
                        line=dict(color=status_color, width=2)),
            name=project.name,
            hovertemplate=(
                f"<b>{project.name}</b><br>"
                f"Start: {project.start_date.strftime('%b %d, %Y')}<br>"
                f"End: {project.end_date.strftime('%b %d, %Y')}<br>"
                f"Status: {project.status}<extra></extra>"
            ),
            showlegend=True,
        ))

        for p_idx, phase in enumerate(project.phases):
            phase_color = PHASE_PALETTE[p_idx % len(PHASE_PALETTE)]
            phase_start_dt = datetime.datetime.combine(phase.start_date, datetime.time())
            phase_days = (phase.end_date - phase.start_date).days

            fig.add_trace(go.Bar(
                x=[phase_days * 86400000],
                y=[project.name],
                base=[phase_start_dt],
                orientation="h",
                marker=dict(color=phase_color, opacity=0.8),
                name=f"{project.name} — {phase.name}",
                hovertemplate=(
                    f"<b>{phase.name}</b><br>"
                    f"{phase.start_date.strftime('%b %d, %Y')} → "
                    f"{phase.end_date.strftime('%b %d, %Y')}<extra></extra>"
                ),
                showlegend=False,
            ))

        for ms in project.milestones:
            ms_dt = datetime.datetime.combine(ms.date, datetime.time())
            ms_color = ms.marker_color(today)
            fig.add_trace(go.Scatter(
                x=[ms_dt],
                y=[project.name],
                mode="markers+text",
                marker=dict(symbol="diamond", size=12, color=ms_color,
                            line=dict(color="black", width=1)),
                text=[ms.name],
                textposition="top center",
                textfont=dict(size=9),
                hovertemplate=(
                    f"<b>{ms.name}</b><br>"
                    f"{ms.date.strftime('%b %d, %Y')}<extra></extra>"
                ),
                showlegend=False,
            ))

    today_dt = datetime.datetime.combine(today, datetime.time())
    fig.add_shape(
        type="line",
        x0=today_dt, x1=today_dt,
        y0=0, y1=1, yref="paper",
        line=dict(color=cfg.TODAY_LINE_COLOR, width=2, dash="dash"),
    )
    fig.add_annotation(
        x=today_dt, y=1.05, yref="paper",
        text=f"Today ({today.strftime('%b %d, %Y')})",
        showarrow=False,
        font=dict(color=cfg.TODAY_LINE_COLOR, size=11, family="Arial Black"),
    )

    for ref in reference_lines:
        ref_dt = datetime.datetime.combine(ref.date, datetime.time())
        fig.add_shape(
            type="line",
            x0=ref_dt, x1=ref_dt,
            y0=0, y1=1, yref="paper",
            line=dict(color=ref.color, width=2, dash="dashdot"),
        )
        fig.add_annotation(
            x=ref_dt, y=1.05, yref="paper",
            text=ref.name,
            showarrow=False,
            font=dict(color=ref.color, size=11, family="Arial Black"),
        )

    fig.update_layout(
        title=title,
        barmode="overlay",
        xaxis=dict(title="Date", type="date"),
        yaxis=dict(title="", autorange="reversed"),
        height=max(400, len(projects) * 120),
        hovermode="closest",
        margin=dict(t=80),
    )

    out = pathlib.Path(html_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.write_html(str(out))
    print(f"✅ Interactive HTML chart saved to {out}")