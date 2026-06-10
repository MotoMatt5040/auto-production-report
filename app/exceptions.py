"""Typed exceptions for the production-report job.

These replace the old bare ``except:`` / ``except Exception as err: raise err`` patterns
so the runner can decide what to do based on *why* something failed.
"""


class ProductionReportError(Exception):
    """Base class for all application-specific errors."""


class ConfigError(ProductionReportError):
    """Configuration is missing or invalid (missing .env key, bad manual_run file,
    missing ODBC driver). Fatal at startup — there is no point retrying."""


class TransientDBError(ProductionReportError):
    """A database call failed in a way that may succeed on retry (timeout, dropped
    connection, deadlock) and did not recover within the retry budget."""


class WorkbookLockedError(ProductionReportError):
    """The target .xlsm could not be opened because it is locked/open elsewhere.
    Non-fatal: the project is carried over to the next run."""


class WorkbookError(ProductionReportError):
    """An Excel/workbook operation failed (missing sheet, macro failure, populate error)."""


class SheetAlreadyExistsError(WorkbookError):
    """The target sheet name (``{project_code} {MMDD}``) already exists — the report
    for this project/date was almost certainly already generated. Logged and skipped,
    NOT carried over."""
