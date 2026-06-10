"""Database access: a cache of pooled SQLAlchemy engines plus a resilient query runner.

Replaces ``DataBaseAccessInfo`` (a mutable singleton that called ``create_engine`` on
every property access and blocked on ``input()``/``webbrowser`` on failure).

* One pooled engine per ``(server, database)`` — reused across the run.
* ``run_query`` retries transient DB errors with exponential backoff, then raises
  :class:`TransientDBError`.
* A missing ODBC driver / bad credentials raises :class:`ConfigError` (fail fast).
"""
from __future__ import annotations

import time
import urllib.parse
from typing import Optional

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import DBAPIError, InterfaceError, OperationalError

from app.config import DbCreds
from app.exceptions import ConfigError, TransientDBError
from utils.logger_config import logger

# pyodbc error substrings that mean "driver/config problem", not "try again".
_FATAL_DRIVER_MARKERS = (
    "data source name not found",
    "im002",
    "can't open lib",
    "driver not found",
    "ODBC Driver",
)


def _is_transient(exc: BaseException) -> bool:
    """Heuristic: connection-level failures are retryable; auth/driver failures are not."""
    msg = str(exc).lower()
    if any(m.lower() in msg for m in _FATAL_DRIVER_MARKERS):
        return False
    if isinstance(exc, (OperationalError, InterfaceError)):
        return True
    if isinstance(exc, DBAPIError):
        return bool(getattr(exc, "connection_invalidated", False)) or "timeout" in msg
    return False


def _is_fatal_config(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return any(m.lower() in msg for m in _FATAL_DRIVER_MARKERS) or "login failed" in msg


class EngineCache:
    """Lazily creates and reuses one pooled engine per ``(server, database)``.

    The ProMark engine is created once. Voxco project databases vary per project, so
    each distinct project DB (and the VoxcoSystem catalog) gets its own cached engine.
    """

    def __init__(self) -> None:
        self._engines: dict[tuple[str, str], Engine] = {}

    def get(self, creds: DbCreds, database: str) -> Engine:
        key = (creds.server, database)
        engine = self._engines.get(key)
        if engine is None:
            logger.debug("Creating pooled engine for %s / %s", creds.server, database)
            engine = create_engine(
                self._url(creds, database),
                pool_pre_ping=True,   # detect stale connections before use
                pool_recycle=1800,    # recycle connections every 30 min
            )
            self._engines[key] = engine
        return engine

    @staticmethod
    def _url(creds: DbCreds, database: str) -> str:
        odbc = (
            f"DRIVER={creds.driver};SERVER={creds.server};DATABASE={database};"
            f"UID={creds.user};PWD={creds.password}"
        )
        return f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(odbc)}"

    def dispose_all(self) -> None:
        for engine in self._engines.values():
            engine.dispose()
        self._engines.clear()


def run_query(
    engine: Engine,
    sql: str,
    params: Optional[dict] = None,
    *,
    attempts: int = 3,
    backoff: float = 2.0,
) -> pd.DataFrame:
    """Execute ``sql`` (with named bind ``params``) and return a DataFrame.

    Retries transient failures up to ``attempts`` times with exponential backoff.
    Raises :class:`ConfigError` on driver/auth problems, :class:`TransientDBError`
    when retries are exhausted.
    """
    statement = text(sql)
    last_exc: Optional[BaseException] = None

    for attempt in range(1, attempts + 1):
        try:
            with engine.connect() as cnxn:
                return pd.read_sql_query(statement, cnxn, params=params or {})
        except Exception as exc:  # noqa: BLE001 - classified below
            last_exc = exc
            if _is_fatal_config(exc):
                logger.error("Database configuration/driver error: %s", exc)
                raise ConfigError(str(exc)) from exc
            if not _is_transient(exc) or attempt == attempts:
                break
            wait = backoff ** (attempt - 1)
            logger.warning(
                "Transient DB error (attempt %d/%d), retrying in %.1fs: %s",
                attempt, attempts, wait, exc,
            )
            time.sleep(wait)

    raise TransientDBError(
        f"Query failed after {attempts} attempt(s): {last_exc}"
    ) from last_exc
