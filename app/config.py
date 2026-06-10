"""Configuration loading and validation.

Secrets and paths come from the ``.env`` file (loaded by ``main`` via python-dotenv).
Ad-hoc / manual runs come from ``manual_run.json`` at the repo root.

All SQL now lives in :mod:`app.queries` — the SQL-fragment env keys that used to live in
``.env`` are no longer read and can be deleted from ``.env``.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from app.exceptions import ConfigError

# Repo root (one level above app/).
ROOT = Path(__file__).resolve().parents[1]
MANUAL_RUN_PATH = ROOT / "manual_run.json"
CARRYOVER_PATH = ROOT / "carryover.json"

DATE_FORMAT = "%Y-%m-%d"


@dataclass(frozen=True)
class DbCreds:
    """Connection parameters for a SQL Server target.

    ``database`` may be overridden per-call (Voxco project DBs vary by project), so it
    is intentionally not part of the credential set used to *reach* the server.
    """
    server: str
    user: str
    password: str
    driver: str = "{ODBC Driver 17 for SQL Server}"


@dataclass(frozen=True)
class Settings:
    src: Path                 # project root, e.g. i:/PROJ/
    planner: Path             # planner/forecast root, e.g. r:/USERS/PLANNER/
    voxco_system_db: str      # the VoxcoSystem catalog name (env: voxco)
    promark_db: str           # ProMark catalog name (env: caligula)
    promark: DbCreds          # coreserver / coreuser / corepassword
    voxco: DbCreds            # cc3server / cc3user / cc3password (db chosen per project)


@dataclass(frozen=True)
class ManualRun:
    project_id: str
    date: date


def _require(key: str) -> str:
    val = os.environ.get(key)
    if val is None or val.strip() == "":
        raise ConfigError(f"Missing required environment variable: {key!r} (check your .env)")
    return val.strip()


def load_settings() -> Settings:
    """Read and validate settings from the environment. Raises :class:`ConfigError`
    if a required key is missing."""
    return Settings(
        src=Path(_require("src")),
        planner=Path(_require("planner")),
        voxco_system_db=_require("voxco"),
        promark_db=_require("caligula"),
        promark=DbCreds(
            server=_require("coreserver"),
            user=_require("coreuser"),
            password=_require("corepassword"),
        ),
        voxco=DbCreds(
            server=_require("cc3server"),
            user=_require("cc3user"),
            password=_require("cc3password"),
        ),
    )


def load_manual_runs(path: Path = MANUAL_RUN_PATH) -> list[ManualRun]:
    """Load ad-hoc run entries from ``manual_run.json``.

    Returns ``[]`` when the file is absent or empty (→ caller uses AUTO mode).

    Schema::

        [
            {"project_id": "13427C", "date": "2026-05-07"},
            {"project_id": "13429",  "date": "2026-05-07"}
        ]
    """
    if not path.exists():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8") or "[]")
    except json.JSONDecodeError as e:
        raise ConfigError(f"{path.name} is not valid JSON: {e}") from e

    if not raw:
        return []
    if not isinstance(raw, list):
        raise ConfigError(f"{path.name} must contain a JSON array of {{project_id, date}} objects.")

    runs: list[ManualRun] = []
    for i, entry in enumerate(raw):
        try:
            pid = str(entry["project_id"]).strip()
            d = datetime.strptime(str(entry["date"]).strip(), DATE_FORMAT).date()
        except (KeyError, TypeError, ValueError) as e:
            raise ConfigError(
                f"{path.name} entry #{i} is invalid (need project_id + date as YYYY-MM-DD): {e}"
            ) from e
        if not pid:
            raise ConfigError(f"{path.name} entry #{i} has an empty project_id.")
        runs.append(ManualRun(project_id=pid, date=d))
    return runs
