"""Carry-over of unfinished projects to the next run.

When a project can't be completed (locked file, transient DB failure, workbook error)
it is persisted to ``carryover.json`` with its **original data date**. The next run
retries it first (for that original date), indefinitely, until it succeeds — so no
report day is ever silently lost.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path

from app.config import CARRYOVER_PATH, DATE_FORMAT
from utils.logger_config import logger


@dataclass
class PendingItem:
    project_id: str
    date: date          # ORIGINAL data date, preserved across retries
    reason: str         # last failure reason (for the log)
    first_failed: date  # when it first failed; retried indefinitely

    def key(self) -> tuple[str, date]:
        return (self.project_id, self.date)


def load_pending(path: Path = CARRYOVER_PATH) -> list[PendingItem]:
    if not path.exists():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8") or "[]")
    except json.JSONDecodeError as e:
        logger.error("carryover file is corrupt (%s); ignoring it this run", e)
        return []

    items: list[PendingItem] = []
    for entry in raw:
        try:
            items.append(
                PendingItem(
                    project_id=str(entry["project_id"]),
                    date=datetime.strptime(entry["date"], DATE_FORMAT).date(),
                    reason=str(entry.get("reason", "")),
                    first_failed=datetime.strptime(
                        entry.get("first_failed", entry["date"]), DATE_FORMAT
                    ).date(),
                )
            )
        except (KeyError, TypeError, ValueError) as e:
            logger.warning("Skipping malformed carryover entry %r: %s", entry, e)
    return items


def save_pending(items: list[PendingItem], path: Path = CARRYOVER_PATH) -> None:
    payload = []
    for it in items:
        d = asdict(it)
        d["date"] = it.date.strftime(DATE_FORMAT)
        d["first_failed"] = it.first_failed.strftime(DATE_FORMAT)
        payload.append(d)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
