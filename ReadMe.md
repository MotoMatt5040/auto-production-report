# Auto Production Report

Daily job that pulls production/sample data from SQL Server (ProMark + Voxco) and fills
the per-project Excel `.xlsm` production reports via xlwings (Excel COM + a VBA macro).

# Installation
## Python
Clone repo and run `pip install -r requirements.txt` to install all necessary packages.
## Exe
Simply download the exe from the dist page and run it. No installation necessary but the .env file will be required to run it.

# Running

```
python main.py
```

### Modes (no code edits to switch)

- **AUTO** (default — for the scheduled daily job): generates a report for every
  project/date row in `tblGPCPHDaily` whose `RecDate` is within the lookback window,
  each project for its own recdate. Runs when there is no `manual_run.json`.
  - `python main.py` — last **1 day** (yesterday's data, the daily job).
  - `python main.py --days N` — backfill the last **N** days, e.g. `--days 60`.
- **MANUAL** (ad-hoc reruns): create `manual_run.json` at the repo root listing the exact
  `(project_id, date)` pairs to (re)generate — see `manual_run.example.json`. Delete/empty
  the file to return to AUTO mode.

### Carry-over

If a project can't be completed (workbook open/locked, transient DB error, etc.) it is
written to `carryover.json` with its **original data date** and retried first on the next
run — indefinitely, until it succeeds. The run logs a summary (succeeded / skipped /
carried-over) and exits non-zero if anything failed or was skipped, so the scheduler notices.

# Configuration

Credentials and paths live in `.env` (not committed): `src`, `planner`, `coreserver`,
`caligula`, `coreuser`, `corepassword`, `cc3server`, `cc3user`, `cc3password`,
`voxco` (= `VoxcoSystem`). SQL now lives in `app/queries.py` as parameterized statements —
the old SQL-fragment keys in `.env` are no longer used and can be removed.

# Layout

| Path | Responsibility |
|------|----------------|
| `main.py` | Entry point; chooses AUTO/MANUAL, applies carry-over, logs summary. |
| `app/config.py` | Loads/validates settings from `.env` and `manual_run.json`. |
| `app/queries.py` | All SQL as named-bind parameterized statements. |
| `app/db.py` | Pooled SQLAlchemy engines (cached per server/db) + retrying query runner. |
| `app/datasource.py` | High-level data fetches (production report, Voxco sample/PREL). |
| `app/workbook.py` | xlwings + VBA wrapper; preserves the exact cell layout. |
| `app/runner.py` | Per-project loop with isolation + carry-over. |
| `app/carryover.py` | Persists/loads unfinished projects (`carryover.json`). |
| `app/exceptions.py` | Typed exceptions used to decide retry/skip/carry-over. |
| `utils/logger_config.py` | Colored console + rotating file logging (`logs/`). |

