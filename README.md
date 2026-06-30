# Outlook Email ETL Pipeline

Extract structured metric data from HTML email reports, store it in a local SQLite database, and export it to Excel — with incremental execution support.

## Architecture

```
 Outlook (COM)
     │
     ▼
 ┌─────────────────────┐
 │   outlook_reader    │  Extract: fetch emails matching keywords
 └─────────┬───────────┘
           │  HTML body
           ▼
 ┌─────────────────────┐
 │   table_parser      │  Transform: parse HTML tables → structured rows
 └─────────┬───────────┘
           │  dicts
           ▼
 ┌─────────────────────┐
 │  sqlite_manager     │  Load: persist to SQLite (dedup via UNIQUE)
 └─────────┬───────────┘
           │
           ├── processed_emails  (incremental control)
           │
           ▼
 ┌─────────────────────┐
 │  excel_exporter     │  Export: pandas → .xlsx per day
 └─────────────────────┘
```

## How It Works

1. **Connect to Outlook** via COM (pywin32) — reads from a PST file or the default MAPI profile
2. **Navigate** to a configurable folder path and filter emails by HTML body keywords
3. **Parse** each matching email's HTML to find the metrics table using header matching
4. **Validate** each row: reject headers, non-metric dimensions, and numeric-only scenarios
5. **Insert** into SQLite with a `UNIQUE` constraint to prevent duplicates across runs
6. **Track** the latest processed email timestamp in `etl_control` for incremental runs
7. **Export** daily Excel files via pandas (local + optional cloud sync)

### Execution Modes

```
PRODUCTION  → Resume from last processed email (incremental)
HISTORICAL  → Process last 90 days (full backfill)
TEST        → Process a small time window for validation
```

## Setup

### Prerequisites
- Windows (required for pywin32 / Outlook COM)
- Microsoft Outlook installed and configured
- Python 3.12+

### Installation

```bash
git clone <repo-url>
cd bes_outlook_extractor
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Configuration

Edit `ruta.py` to set your environment:

```python
DATA_OUTLOOK = None  # or Path("C:/path/to/archive.pst")
FOLDER_PATH = ...    # in outlook_reader.py
REPORT_KEYWORDS = ...  # in outlook_reader.py
```

## Usage

```bash
# Production (incremental — resumes from last run)
python main.py

# Historical backfill
python main.py --mode HISTORICAL

# Quick test
python main.py --mode TEST --days 1
```

## Project Structure

```
bes_outlook_extractor/
├── main.py                      Orchestrator (entry point)
├── ruta.py                      Configuration paths
├── requirements.txt             Dependencies
├── README.md                    This file
│
├── src/
│   ├── extract/
│   │   └── outlook_reader.py    Outlook COM extraction
│   ├── discover/
│   │   └── table_parser.py      HTML table parsing
│   ├── database/
│   │   └── sqlite_manager.py    SQLite persistence
│   ├── control/
│   │   └── processed_emails.py  Incremental execution tracking
│   └── export/
│       └── excel_exporter.py    Excel export
│
└── data/
    ├── database/bes.db          SQLite database
    ├── processed/               Exported Excel files
    └── raw/html/                Optional HTML backups
```

## Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| **SQLite** | Serverless, single file, supports SQL queries and UNIQUE constraints |
| **COM/pywin32** | Only way to access full Outlook HTML body (no REST API for on-prem) |
| **BeautifulSoup + lxml** | Fast, tolerant of malformed HTML common in auto-generated emails |
| **6-column header match** | Finds the correct table among many in a multi-table HTML email |
| **Scenario validation** | Filters out numeric-only rows from non-metric tables (Business, Resource) |
| **EntryID validation** | Rejects timestamp-like IDs that would cause false dedup |
| **Incremental tracking** | Only re-processes the latest failed/new emails on each run |
| **UNIQUE constraint** | Double safety: code check + database constraint |

## Database Schema

**`bes_status`** — stores all parsed metric rows

```sql
UNIQUE(entry_id, dimension, scenario, received_time)
```

**`etl_control`** — tracks the latest processed timestamp for incremental runs

```sql
last_email_received  -- checkpoint for next PRODUCTION run
```

## Dependencies

- `beautifulsoup4` + `lxml` — HTML parsing
- `pandas` + `openpyxl` — Excel export
- `pywin32` — Outlook COM interface
