"""Per-project orchestration with isolation and carry-over.

Rewrite of ``looprunner``. Differences from the old loop:

* No bare ``except:`` around everything — each project runs in its own try/except, so
  one failure never aborts the whole daily run.
* No blocking ``input()`` on a locked file — it raises and the project is carried over.
* Files are still opened once per 5-digit project number and saved/closed after the
  last variant, exactly as before. We sort the work so a project's variants stay
  adjacent and the non-``C`` (LL) variant is processed before its ``C`` variant.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from app.carryover import PendingItem
from app.datasource import DataSource
from app.exceptions import (
    SheetAlreadyExistsError,
    TransientDBError,
    WorkbookError,
    WorkbookLockedError,
)
from app.workbook import WorkbookHandler
from utils.logger_config import logger

# A unit of work: produce the report for one project code on one data date.
WorkItem = tuple[str, date]


@dataclass
class RunSummary:
    succeeded: list[WorkItem] = field(default_factory=list)
    skipped: list[tuple[WorkItem, str]] = field(default_factory=list)   # already-generated, etc.
    failed: list[tuple[WorkItem, str]] = field(default_factory=list)    # carried over

    def as_pending(self, today: date, prior: dict[WorkItem, PendingItem]) -> list[PendingItem]:
        """Build the new carry-over list from this run's failures, preserving the
        original ``first_failed`` date for items that were already pending."""
        pending: list[PendingItem] = []
        for (pid, d), reason in self.failed:
            prev = prior.get((pid, d))
            pending.append(
                PendingItem(
                    project_id=pid,
                    date=d,
                    reason=reason,
                    first_failed=prev.first_failed if prev else today,
                )
            )
        return pending


def _sort_key(item: WorkItem):
    """Group by 5-digit project number + date, LL ('') before C variant."""
    pid, d = item
    base = pid[:5]
    suffix = pid[5:].upper()  # '' for LL, 'C' for cell
    return (base, d, suffix)


def run(run_list: list[WorkItem], ds: DataSource) -> RunSummary:
    """Process every work item with per-project isolation. Returns a RunSummary."""
    summary = RunSummary()
    work = sorted(set(run_list), key=_sort_key)
    if not work:
        logger.info("Nothing to process.")
        return summary

    wh = WorkbookHandler().start()
    try:
        open_group: str | None = None  # the 5-digit project number whose file is open
        for index, (project, recdate) in enumerate(work):
            project_number = project[:5]
            try:
                # Open the file once per project number.
                if open_group != project_number:
                    wh.project_id = project_number
                    wh.project_code = project
                    wh.path = Path(_report_path(project_number))
                    wh.check_path()
                    wh.set_workbook()
                    open_group = project_number

                wh.project_code = project
                wh.date = recdate
                _build_report(wh, ds, project, recdate)

                summary.succeeded.append((project, recdate))
                logger.info("Completed - %s - %s", project, recdate)

            except SheetAlreadyExistsError as e:
                logger.warning("Skipping %s %s - %s", project, recdate, e)
                summary.skipped.append(((project, recdate), str(e)))
            except WorkbookLockedError as e:
                logger.warning("Carrying over %s %s - locked: %s", project, recdate, e)
                summary.failed.append(((project, recdate), f"locked: {e}"))
                open_group = None  # couldn't open; force re-open next group
            except (TransientDBError, WorkbookError) as e:
                logger.error("Carrying over %s %s - %s", project, recdate, e, exc_info=True)
                summary.failed.append(((project, recdate), str(e)))
            except Exception as e:  # noqa: BLE001 - unexpected; isolate and continue
                logger.exception("Carrying over %s %s - unexpected error", project, recdate)
                summary.failed.append(((project, recdate), repr(e)))
            finally:
                # Save & close after the last variant of this project number.
                is_last = index == len(work) - 1 or work[index + 1][0][:5] != project_number
                if is_last and open_group == project_number and wh.workbook is not None:
                    try:
                        wh.save()
                        wh.close()
                    except Exception as e:  # noqa: BLE001
                        logger.error("Failed saving/closing %s: %s", project_number, e)
                    open_group = None
    finally:
        wh.quit()

    return summary


def _report_path(project_number: str) -> str:
    import os
    return str(
        Path(os.environ['src']) / project_number / 'PRODUCTION'
        / f"{project_number}_Production_Report.xlsm"
    )


def _build_report(wh: WorkbookHandler, ds: DataSource, project: str, recdate: date) -> None:
    """The per-variant work previously in looprunner.read_excel()."""
    wh.copy_sheet()
    if wh.project_code.upper().endswith("C"):
        wh.active_sheet = "PROJ#C DATE (2)"
    else:
        wh.active_sheet = "PROJ# DATE (2)"
    wh.set_active_sheet_name()

    wh.data = ds.production_report(wh.project_code, recdate)
    wh.populate_all()
    wh.populate_expected_loi()

    voxco_db = ds.voxco_project_database(wh.project_code)
    sample = ds.voxco_sample(voxco_db, recdate)
    prel = ds.voxco_prel(voxco_db, recdate)
    wh.perf_swap()
    wh.populate_perf(sample_data=sample, prel_data=prel)
