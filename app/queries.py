"""All SQL for the production-report job, as parameterized statements.

This replaces the old ``SQLDictionary`` + ``.env`` fragment concatenation. Two rules:

* **Values** (project ids, dates, names) are passed as named bind parameters
  (``:projectid``, ``:recdate``, ...). They are never string-formatted into the SQL,
  so there is no injection surface.
* **Identifiers** that must be dynamic — only the Voxco per-project *database name* —
  cannot be bind parameters in T-SQL. They are validated against a strict allowlist
  (:func:`_validate_voxco_db`) before being formatted into the ``[db].[dbo].[table]``
  references.

The query *text* mirrors the original concatenated SQL so results are unchanged.
"""
from __future__ import annotations

import re

from app.exceptions import ConfigError

# ---------------------------------------------------------------------------
# ProMark queries — pure bind parameters.
# ---------------------------------------------------------------------------

ACTIVE_PROJECT_IDS = """
SELECT DISTINCT projectid, recdate
FROM tblGPCPHDaily
WHERE RecDate >= :recdate
"""

PRODUCTION_REPORT = """
SELECT DISTINCT eid, refname, recloc, tenure, hrs,
       SUM(connecttime) AS connecttime, SUM(pausetime) AS pausetime,
       cms, intal, mph, totaldials
FROM tblProduction
INNER JOIN tblEmployees   ON empid = eid
INNER JOIN tblAspenProdII ON tblAspenProdII.empid = tblProduction.eid
                         AND tblAspenProdII.projectid = tblProduction.projectid
                         AND tblAspenProdII.recdate = tblProduction.recdate
WHERE tblProduction.projectid = :projectid
  AND tblProduction.recdate = :recdate
GROUP BY eid, refname, tblProduction.recloc, tenure, hrs,
         cms, cph, mph, dpc, dph, totaldials, intal
"""

PRODUCTION_DISPO = """
SELECT inc, mean, projname
FROM tbldispo
WHERE projid = :projectid
  AND dispodate = :recdate
  AND SingleDay = 0
"""

PRODUCTION_AVG_LENGTH = """
SELECT avgcph, avgmph, avglen
FROM tblBlueBookProjMaster
WHERE projectid = :projectid
  AND recdate = :recdate
"""

# Looks up the Voxco project database NAME for a project number. Run against VoxcoSystem.
VOXCO_PROJECT_DATABASE = """
SELECT 'Voxco_Project_' + CAST(k_Id AS nvarchar(255)) AS ProjectDatabase
FROM [VoxcoSystem].[dbo].[tblObjects]
WHERE Type = 2
  AND name = :project_number
"""

# ---------------------------------------------------------------------------
# Voxco per-project queries — database name is a validated identifier.
# Time bounds (:start / :end) remain bind parameters.
# ---------------------------------------------------------------------------

_VOXCO_DB_RE = re.compile(r"^Voxco_Project_\d+$")


def _validate_voxco_db(database: str) -> str:
    """Allowlist-check a Voxco project database name before interpolating it as an
    identifier. Prevents injection via the (necessarily) formatted db name."""
    if not isinstance(database, str) or not _VOXCO_DB_RE.match(database):
        raise ConfigError(f"Refusing to use unexpected Voxco database name: {database!r}")
    return database


def voxco_scanned_sql(database: str) -> str:
    """Scanned sample data. Params: ``start``, ``end`` (datetimes/strings)."""
    db = _validate_voxco_db(database)
    return f"""
SELECT HisCallNumber, HisCaseResult
FROM [{db}].[dbo].[Historic]
WHERE HisCallDate >= :start
  AND HisCallDate <  :end
  AND HisCaseResult IS NOT NULL
"""


def voxco_prel_sql(database: str) -> str:
    """PREL sample data (scanned). Params: ``start``, ``end``."""
    db = _validate_voxco_db(database)
    return f"""
SELECT RpsContent, HisCaseResult
FROM [{db}].[dbo].[Historic]
INNER JOIN [{db}].[dbo].[Response]
        ON [Response].[RpsRespondent] = [Historic].[HisRespondent]
WHERE HisCallDate >= :start
  AND HisCallDate <  :end
  AND RpsQuestion = 'PREL'
"""


def voxco_unscanned_sql(database: str) -> str:
    """Unscanned sample data. Params: ``start``, ``end``."""
    db = _validate_voxco_db(database)
    return f"""
SELECT HisRespondent, RpsContent, HisCaseResult
FROM [{db}].[dbo].[Historic] AS Historic
LEFT JOIN [{db}].[dbo].[Response] AS Response
       ON [Response].[RpsRespondent] = [Historic].[HisRespondent]
      AND Response.RpsQuestion = 'PREL'
WHERE (Response.RpsRespondent IS NULL OR Response.RpsContent = 9)
  AND A4SCallNumber IS NULL
  AND HisCaseResult IS NOT NULL
  AND HisCallDate >= :start
  AND HisCallDate <  :end
"""
