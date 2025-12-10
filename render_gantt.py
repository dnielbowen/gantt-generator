#!/usr/bin/env python3
"""Render a Plotly Gantt chart for a Microsoft Planner export (CSV/XLSX)."""
from __future__ import annotations

import argparse
import hashlib
import re
from datetime import datetime, timedelta
from fnmatch import fnmatchcase
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
from plotly.graph_objs import Figure

# Avatar color palette for assignees
AVATAR_COLORS = [
    "#E57373", "#81C784", "#64B5F6", "#FFD54F", "#BA68C8",
    "#4DB6AC", "#FF8A65", "#A1887F", "#90A4AE", "#F06292",
    "#AED581", "#7986CB", "#FFB74D", "#4DD0E1", "#9575CD",
]

DATE_COLUMNS: Iterable[str] = (
    "Created Date",
    "Start date",
    "Due date",
    "Completed Date",
)
PROGRESS_TO_PERCENT = {
    "Not started": 0,
    "In progress": 50,
    "Complete": 100,
    "Completed": 100,
}
DEFAULT_DURATION_DAYS = 7


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        "--csv",
        dest="source",
        type=Path,
        default=Path("input.csv"),
        help="Path to the Planner export (CSV or XLSX)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("gantt.html"),
        help="Destination HTML file for the Plotly chart",
    )
    parser.add_argument(
        "--title",
        type=str,
        default=None,
        help="Overrides the chart title (defaults to the input filename)",
    )
    parser.add_argument(
        "--exclude-bucket",
        dest="exclude_buckets",
        action="append",
        default=[],
        metavar="PATTERN",
        help=(
            "Glob pattern for bucket names to exclude; repeat for multiple "
            "patterns (e.g., --exclude-bucket '*4.1*')"
        ),
    )
    return parser.parse_args()


def load_tasks(source_path: Path) -> pd.DataFrame:
    if not source_path.exists():
        raise FileNotFoundError(f"Input not found: {source_path}")

    df = _read_planner_export(source_path)
    df.rename(columns=lambda col: col.strip(), inplace=True)

    for column in DATE_COLUMNS:
        if column in df.columns:
            df[column] = pd.to_datetime(
                df[column], format="%m/%d/%Y", errors="coerce"
            )

    schedule = df.apply(_derive_schedule, axis=1)
    df = pd.concat([df, schedule], axis=1)
    df = df.dropna(subset=["Start", "Finish"]).copy()

    df["Duration (days)"] = (df["Finish"] - df["Start"]).dt.days + 1
    df["Progress %"] = (
        df.get("Progress", pd.Series(dtype=str))
        .str.strip()
        .str.title()
        .map(PROGRESS_TO_PERCENT)
        .fillna(0)
    )
    df["Late Flag"] = df.get("Late", pd.Series(dtype=str)).astype(str).str.lower() == "true"

    df.sort_values(by=["Start", "Finish", "Task Name"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def _read_planner_export(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, engine="python")
    if suffix in {".xlsx", ".xls"}:
        try:
            return pd.read_excel(path, sheet_name="Tasks")
        except ValueError as exc:  # sheet not found
            raise ValueError(
                f"Worksheet 'Tasks' not found in {path.name}"
            ) from exc
    raise ValueError(f"Unsupported file type: {path.suffix or 'unknown'}")


def _get_initials(name: str) -> str:
    """Extract two-letter initials from a name (first + last, ignoring middle and suffixes)."""
    # Strip common suffixes (normalized: lowercase, no punctuation)
    suffixes = {"phd", "md", "jr", "sr", "ii", "iii", "iv", "esq", "dds", "dvm"}
    parts = name.strip().split()
    # Remove trailing suffix parts (normalize by removing all punctuation)
    while parts and re.sub(r"[^a-z]", "", parts[-1].lower()) in suffixes:
        parts.pop()
    if not parts:
        return "??"
    if len(parts) == 1:
        return parts[0][:2].upper()
    # First and last name, ignoring middle names/initials
    return (parts[0][0] + parts[-1][0]).upper()


def _parse_assignees(assignee_str: str | float) -> list[str]:
    """Parse semicolon-delimited assignee string into list of names."""
    if pd.isna(assignee_str) or not str(assignee_str).strip():
        return []
    return [name.strip() for name in str(assignee_str).split(";") if name.strip()]


def _get_assignee_color(name: str, color_map: dict[str, str]) -> str:
    """Get consistent color for an assignee."""
    if name not in color_map:
        idx = len(color_map) % len(AVATAR_COLORS)
        color_map[name] = AVATAR_COLORS[idx]
    return color_map[name]


def _derive_schedule(row: pd.Series) -> pd.Series:
    start = row.get("Start date")
    finish = row.get("Due date")

    created = row.get("Created Date")
    completed = row.get("Completed Date")

    if pd.isna(start) and not pd.isna(created):
        start = created
    if pd.isna(finish) and not pd.isna(completed):
        finish = completed

    if pd.isna(start) and not pd.isna(finish):
        start = finish - timedelta(days=DEFAULT_DURATION_DAYS)
    if pd.isna(finish) and not pd.isna(start):
        finish = start + timedelta(days=DEFAULT_DURATION_DAYS)

    if not pd.isna(start) and not pd.isna(finish) and finish < start:
        finish = start

    return pd.Series({"Start": start, "Finish": finish})


def build_figure(df: pd.DataFrame, chart_title: str) -> Figure:
    fig = px.timeline(
        df,
        x_start="Start",
        x_end="Finish",
        y="Task Name",
        color="Bucket Name",
        hover_data={
            "Bucket Name": True,
            "Assigned To": True,
            "Priority": True,
            "Progress": True,
            "Progress %": True,
            "Duration (days)": True,
            "Start": True,
            "Finish": True,
            "Late Flag": True,
        },
    )

    fig.update_yaxes(autorange="reversed")

    today = datetime.now()
    fig.add_shape(
        type="line",
        x0=today,
        x1=today,
        yref="paper",
        y0=0,
        y1=1,
        line=dict(color="red", dash="dot", width=2),
    )
    fig.add_annotation(
        x=today,
        y=1,
        yref="paper",
        text="Today",
        showarrow=False,
        font=dict(color="red"),
        xanchor="left",
        yanchor="bottom",
    )

    fig.update_layout(
        title=chart_title,
        xaxis_title="Schedule",
        yaxis_title="Tasks",
        legend_title="Bucket",
        template="plotly_white",
        bargap=0.2,
        hoverlabel=dict(align="left"),
        height=max(600, 40 * len(df) + 200),
        margin=dict(l=240, r=80, t=80, b=40),
    )

    # Add green border to completed tasks
    _style_completed_bars(fig, df)

    # Add avatar icons for assignees
    _add_assignee_avatars(fig, df)

    return fig


def _style_completed_bars(fig: Figure, df: pd.DataFrame) -> None:
    """Add green border to completed task bars."""
    if "Progress %" not in df.columns:
        return

    # Build lookup: task_name -> is_completed
    completed_tasks = set(df.loc[df["Progress %"] == 100, "Task Name"])

    for trace in fig.data:
        if trace.type != "bar":
            continue
        # Each bar trace may contain multiple tasks; style each bar individually
        if trace.y is None:
            continue
        task_names = list(trace.y)
        border_colors = [
            "#00FF00" if task in completed_tasks else "rgba(0,0,0,0)"
            for task in task_names
        ]
        border_widths = [
            2 if task in completed_tasks else 0
            for task in task_names
        ]
        trace.marker.line = dict(color=border_colors, width=border_widths)


def _add_assignee_avatars(fig: Figure, df: pd.DataFrame) -> None:
    """Add avatar circles with initials for each task's assignees."""
    if "Assigned To" not in df.columns:
        return

    color_map: dict[str, str] = {}
    task_names = df["Task Name"].tolist()

    # Calculate spacing based on typical task duration
    avg_duration = (df["Finish"] - df["Start"]).mean()
    spacing = avg_duration * 0.08

    for idx, row in df.iterrows():
        assignees = _parse_assignees(row.get("Assigned To"))
        if not assignees:
            continue

        start_ts = row["Start"]
        task_name = row["Task Name"]

        for i, assignee in enumerate(assignees):
            initials = _get_initials(assignee)
            color = _get_assignee_color(assignee, color_map)

            # Position circles starting from left of bar, spaced apart
            x_pos = start_ts + spacing * (1.0 + i * 2.0)

            # Add circular marker (maintains aspect ratio)
            fig.add_trace(
                dict(
                    type="scatter",
                    x=[x_pos],
                    y=[task_name],
                    mode="markers+text",
                    marker=dict(
                        size=24,
                        color=color,
                        line=dict(color="white", width=2),
                    ),
                    text=initials,
                    textfont=dict(color="white", size=9),
                    textposition="middle center",
                    hovertext=assignee,
                    hoverinfo="text",
                    showlegend=False,
                )
            )


def exclude_buckets(df: pd.DataFrame, patterns: Iterable[str]) -> pd.DataFrame:
    if not patterns or "Bucket Name" not in df.columns:
        return df

    drop_mask = df["Bucket Name"].fillna("").apply(
        lambda bucket: any(fnmatchcase(bucket, pattern) for pattern in patterns)
    )
    return df.loc[~drop_mask].copy()


def main() -> None:
    args = parse_args()
    df = load_tasks(args.source)
    df = exclude_buckets(df, args.exclude_buckets)
    if df.empty:
        raise SystemExit("No tasks with schedule info found in Planner export")

    chart_title = args.title or args.source.stem or "Planner Tasks Timeline"
    fig = build_figure(df, chart_title)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    fig.write_html(args.output, include_plotlyjs="cdn", full_html=True)
    print(f"Wrote Gantt chart to {args.output.resolve()}")


if __name__ == "__main__":
    main()
