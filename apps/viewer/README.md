# Asprova Schedule Viewer

A Flask web application for viewing and browsing Asprova production schedules.

## Features
- Import production schedule CSVs exported from Asprova
- Auto-detect column mappings (flexible CSV format support)
- SQLite database storage
- Daily and weekly schedule list views
- Interactive Gantt chart per machine/resource
- Machine filter and date navigation
- Sample CSV download for testing

## Setup

```bash
pip install flask werkzeug

# Run the app
python app.py
```

Then open: http://localhost:5000

## CSV Format

The importer auto-detects columns by keyword matching. Supported fields:

| Field | Detected Keywords |
|-------|------------------|
| order_id | order, lot, job, work order |
| item_name | item_name, product, description |
| machine_name | resource, machine, equipment |
| start_time | start, begin, from, planned_start |
| end_time | end, finish, to, planned_end |
| process_name | process, operation, activity |
| quantity | qty, quantity, amount |
| status | status, state, condition |

### Supported Date Formats
- `YYYY/MM/DD HH:MM:SS`
- `YYYY-MM-DD HH:MM:SS`
- `YYYY/MM/DD HH:MM`
- `YYYY-MM-DD`
- `MM/DD/YYYY HH:MM:SS`

### Supported Delimiters
Comma (`,`), Tab (`\t`), Semicolon (`;`), Pipe (`|`)

## Pages

- `/` — Dashboard with summary stats
- `/upload` — CSV import page
- `/schedule` — List view (daily/weekly)
- `/gantt` — Gantt chart view
- `/sample_csv` — Download sample CSV
