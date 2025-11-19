
# Component 4 Gantt Renderer

This repository turns the MS Teams task plan into an interactive Plotly Gantt chart. It accepts either the Microsoft Planner CSV export or the equivalent XLSX workbook (the `Tasks` worksheet) and produces a standalone HTML timeline.

## Getting Started

1. **Create/activate the virtual environment**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```
2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
3. **Render the chart**
   ```bash
   .venv/bin/python render_gantt.py --input data/exportedTasks.xlsx --output gantt.html
   ```

## Planner CSV/XLSX Format

Both formats share the same tabular schema; the XLSX version stores it inside a worksheet named `Tasks`. The renderer expects the following columns (extras are ignored):

| Column | Description |
| ------ | ----------- |
| `Task ID` | Unique Planner identifier (free text) |
| `Task Name` | Human-readable title shown on the chart |
| `Bucket Name` | Used to colorize bars |
| `Progress` | One of `Not started`, `In progress`, `Complete/Completed`; converted to a percentage for hover text |
| `Priority` | Free-form priority label |
| `Assigned To` / `Created By` | Displayed in the hover card |
| `Created Date` | Fallback start date when the explicit `Start date` field is empty |
| `Start date` | Preferred schedule start (MM/DD/YYYY) |
| `Due date` | Preferred schedule end (MM/DD/YYYY) |
| `Completed Date` | Used when `Due date` is missing |
| `Late` | Boolean flag (`true`/`false`) surfaced in hover details |
| `Description` | Long-form notes; not displayed but preserved for potential extension |

### Handling Missing Dates

- If `Start date` is blank but `Created Date` exists, the renderer uses the created date.
- If `Due date` is blank but `Completed Date` exists, it uses the completed date.
- If only one side of the window is available, the code synthesizes the other side by adding/subtracting the default 7-day duration.
- Tasks missing both start and finish are dropped because they cannot be charted.
