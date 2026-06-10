"""High-level data fetches for the production report.

Rewrite of ``DataPuller``: no ``try/except: raise`` wrappers, parameterized SQL via
:mod:`app.queries`, and pooled engines via :class:`app.db.EngineCache`. The in-pandas
bucketing for Voxco sample/PREL data is preserved verbatim — it is intentionally faster
than issuing multiple SQL queries.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta

import pandas as pd

from app import queries
from app.config import Settings
from app.db import EngineCache, run_query
from utils.logger_config import logger

# Voxco "day" runs 10:00 -> 10:00 next day (matches the original "'{date} 10:00'" bounds).
_VOXCO_DAY_START = "10:00"

VALID_CASE_RESULTS = [
    "OK", "CO", "05", "06", "07", "10", "12", "13", "14", "16", "22", "23", "25",
    "26", "80", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99",
]


def _voxco_bounds(d) -> dict[str, str]:
    """Build the ``start``/``end`` bind params for a Voxco day window."""
    if isinstance(d, datetime):
        day = d.date()
    elif isinstance(d, date):
        day = d
    else:
        day = pd.to_datetime(d).date()
    start = f"{day} {_VOXCO_DAY_START}"
    end = f"{day + timedelta(days=1)} {_VOXCO_DAY_START}"
    return {"start": start, "end": end}


class DataSource:
    def __init__(self, settings: Settings, engines: EngineCache):
        self._settings = settings
        self._engines = engines

    # -- engine helpers -----------------------------------------------------
    @property
    def _promark_engine(self):
        return self._engines.get(self._settings.promark, self._settings.promark_db)

    @property
    def _voxco_system_engine(self):
        return self._engines.get(self._settings.voxco, self._settings.voxco_system_db)

    def _voxco_project_engine(self, database: str):
        return self._engines.get(self._settings.voxco, database)

    # -- ProMark ------------------------------------------------------------
    def active_projects(self, recdate: date) -> pd.DataFrame:
        """Active project ids/recdates with ``RecDate >= recdate`` (AUTO mode passes
        yesterday)."""
        return run_query(self._promark_engine, queries.ACTIVE_PROJECT_IDS, {"recdate": recdate})

    def production_report(
        self, projectid: str, recdate
    ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Returns (production report, dispo, daily averages) — same triple as before."""
        params = {"projectid": projectid, "recdate": recdate}
        eng = self._promark_engine
        report = run_query(eng, queries.PRODUCTION_REPORT, params)
        dispo = run_query(eng, queries.PRODUCTION_DISPO, params)
        avg = run_query(eng, queries.PRODUCTION_AVG_LENGTH, params)
        return report, dispo, avg

    # -- Voxco --------------------------------------------------------------
    def voxco_project_database(self, project_number: str) -> str:
        """Resolve the Voxco project database NAME for a project number."""
        df = run_query(
            self._voxco_system_engine,
            queries.VOXCO_PROJECT_DATABASE,
            {"project_number": project_number},
        )
        return df["ProjectDatabase"][0]

    def voxco_sample(self, database: str, d) -> dict:
        """Used / live-connect / CO call counts bucketed by HisCallNumber (1..6).

        Pulls one DataFrame and filters in pandas (faster than multiple queries).
        """
        bounds = _voxco_bounds(d)
        df = run_query(self._voxco_project_engine(database), queries.voxco_scanned_sql(database), bounds)

        buckets = {f"used_sample_call_count_{n}": df[df["HisCallNumber"] == n] for n in (1, 2, 3, 4, 5)}
        buckets["used_sample_call_count_6"] = df[df["HisCallNumber"] >= 6]

        live = {
            key.replace("used_sample_call_count_", "live_connects_call_count_"):
                value[value["HisCaseResult"].isin(VALID_CASE_RESULTS)]
            for key, value in buckets.items()
        }
        co = {
            key.replace("used_sample_call_count_", "co_case_count_"):
                value[value["HisCaseResult"] == "CO"]
            for key, value in buckets.items()
        }

        return {
            "used_sample_call_count": {n: buckets[f"used_sample_call_count_{n}"].shape[0] for n in range(1, 7)},
            "live_connects_call_count": {n: live[f"live_connects_call_count_{n}"].shape[0] for n in range(1, 7)},
            "co_case_count": {n: co[f"co_case_count_{n}"].shape[0] for n in range(1, 7)},
        }

    def voxco_prel(self, database: str, d) -> dict:
        """PREL totals + CO counts, aggregated into the ``<>`` / ``0`` / ``1`` groups."""
        bounds = _voxco_bounds(d)
        eng = self._voxco_project_engine(database)
        df = run_query(eng, queries.voxco_prel_sql(database), bounds)
        unscanned = run_query(eng, queries.voxco_unscanned_sql(database), bounds)

        dfs = {"prel_<>": unscanned[["RpsContent", "HisCaseResult"]]}
        dfs.update({f"prel_{n}": df[df["RpsContent"] == str(n)] for n in (0, 1, 2, 3, 4, 5)})

        dfco = {
            key.replace("prel_", "co_"): value[value["HisCaseResult"] == "CO"]
            for key, value in dfs.items()
        }

        return {
            "total": {
                "<>": dfs["prel_<>"].shape[0],
                "0": sum(dfs[k].shape[0] for k in ("prel_0", "prel_1")),
                "1": sum(dfs[k].shape[0] for k in ("prel_2", "prel_3", "prel_4", "prel_5")),
            },
            "co": {
                "<>": dfco["co_<>"].shape[0],
                "0": sum(dfco[k].shape[0] for k in ("co_0", "co_1")),
                "1": sum(dfco[k].shape[0] for k in ("co_2", "co_3", "co_4", "co_5")),
            },
        }
