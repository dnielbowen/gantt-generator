#!/usr/bin/env python3
"""Render a Plotly Gantt chart for a Microsoft Planner export (CSV/XLSX)."""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
from plotly.graph_objs import Figure

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

    return fig


def main() -> None:
    args = parse_args()
    df = load_tasks(args.source)
    if df.empty:
        raise SystemExit("No tasks with schedule info found in Planner export")

    chart_title = args.title or args.source.stem or "Planner Tasks Timeline"
    fig = build_figure(df, chart_title)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    fig.write_html(args.output, include_plotlyjs="cdn", full_html=True)
    print(f"Wrote Gantt chart to {args.output.resolve()}")


if __name__ == "__main__":
    main()
