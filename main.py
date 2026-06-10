"""Entry point for the auto production report job.

Run modes (no code edits needed to switch):

* AUTO (default, for the daily scheduled job): generate reports for every project/date
  row in ``tblGPCPHDaily`` with ``RecDate`` within the lookback window. Each project is
  reported for its OWN recdate. Used when ``manual_run.json`` is absent or empty.
  - ``python main.py``            -> last 1 day (the daily job)
  - ``python main.py --days 60``  -> backfill the last 60 days
* MANUAL: process the explicit (project_id, date) pairs listed in ``manual_run.json``.

Either way, projects that failed or were locked on the previous run (``carryover.json``)
are retried first, for their original dates, before the new work.
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

from app import carryover, config
from app.datasource import DataSource
from app.db import EngineCache
from app.exceptions import ConfigError
from app.runner import RunSummary, WorkItem, run
from utils.logger_config import logger


def _build_run_list(
    ds: DataSource, pending: list[carryover.PendingItem], lookback_days: int
) -> list[WorkItem]:
    """Carry-over items (original dates) first, then today's work, de-duped."""
    items: list[WorkItem] = [(p.project_id, p.date) for p in pending]

    manual = config.load_manual_runs()
    if manual:
        logger.info("MANUAL mode: %d entr(ies) from %s", len(manual), config.MANUAL_RUN_PATH.name)
        items += [(m.project_id, m.date) for m in manual]
    else:
        since = date.today() - timedelta(days=lookback_days)
        df = ds.active_projects(since)
        logger.info(
            "AUTO mode: %d project/date row(s) with RecDate >= %s (lookback %d day(s))",
            df.shape[0], since, lookback_days,
        )
        # Pair each project with ITS OWN recdate from the table (not a single date).
        for pid, recdate in zip(df["projectid"], df["recdate"]):
            items.append((str(pid), pd.to_datetime(recdate).date()))

    # De-dupe on (project, date), preserving order (carry-over first).
    seen: set[WorkItem] = set()
    deduped: list[WorkItem] = []
    for it in items:
        if it not in seen:
            seen.add(it)
            deduped.append(it)
    return deduped


def _log_summary(summary: RunSummary, carried: int) -> None:
    logger.info(
        "Run complete: %d succeeded, %d skipped, %d carried over (%d total pending).",
        len(summary.succeeded), len(summary.skipped), len(summary.failed), carried,
    )
    for item, reason in summary.failed:
        logger.warning("  PENDING %s %s - %s", item[0], item[1], reason)
    for item, reason in summary.skipped:
        logger.info("  SKIPPED %s %s - %s", item[0], item[1], reason)


def _parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate production reports.")
    parser.add_argument(
        "--days", type=int, default=1, metavar="N",
        help="AUTO mode lookback window: include project/date rows with RecDate within "
             "the last N days. Default 1 (yesterday's data, the daily job). Ignored when "
             "manual_run.json is present. Example: --days 60 to backfill 60 days.",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = _parse_args(argv)
    if args.days < 1:
        logger.critical("--days must be >= 1 (got %s)", args.days)
        return 2

    logger.info("STARTING")
    try:
        settings = config.load_settings()
    except ConfigError as e:
        logger.critical("Configuration error - aborting: %s", e)
        return 2

    engines = EngineCache()
    ds = DataSource(settings, engines)

    try:
        pending = carryover.load_pending()
        prior = {p.key(): p for p in pending}

        run_list = _build_run_list(ds, pending, args.days)
        summary = run(run_list, ds)

        # Recompute carry-over: clear successes/skips, keep+append failures.
        new_pending = summary.as_pending(date.today(), prior)
        carryover.save_pending(new_pending)
        _log_summary(summary, len(new_pending))

    except ConfigError as e:
        logger.critical("Configuration error during run - aborting: %s", e)
        return 2
    finally:
        engines.dispose_all()

    # Non-zero exit if anything failed/carried over or was skipped, so the scheduler notices.
    return 1 if (summary.failed or summary.skipped) else 0


if __name__ == "__main__":
    sys.exit(main())
