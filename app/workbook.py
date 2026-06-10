"""Excel workbook handling via xlwings (Excel COM + the BLANK_Production VBA macro).

Rewrite of ``WorkbookHandler`` that PRESERVES the Excel behavior exactly — same VBA
macro call, same cell layout, same fill-down formulas, same perf-sheet row scans. The
only changes are robustness:

* The ``xw.App`` is created lazily (``start()`` / context manager), not at import time,
  and is always quit in a ``finally`` by the runner.
* Opening a locked workbook raises :class:`WorkbookLockedError` instead of blocking on
  ``input()`` — the runner carries the project over to the next run.
* A duplicate sheet name raises :class:`SheetAlreadyExistsError` (was a malformed
  ``raise (e, ...)``); the runner logs it and skips.
"""
from __future__ import annotations

import os
import shutil
from datetime import date, timedelta
from pathlib import Path
from typing import Union

import numpy as np
import pandas as pd
import xlwings as xw

from app.exceptions import SheetAlreadyExistsError, WorkbookError, WorkbookLockedError
from utils.logger_config import logger


class WorkbookHandler:
    """Manages one Excel application instance and the workbook currently open in it.

    Lifecycle: ``start()`` (or use as a context manager) → set project/date → open →
    populate → ``save()`` / ``close()`` → ``quit()``.
    """

    def __init__(self):
        self.app: xw.App | None = None
        self._projectid = None
        self._projectCode = None
        self._path = None
        self._wb = None
        self._activeSheet = None
        self._activeSheetName = None
        self._productionData = pd.DataFrame
        self._date = date.today() - timedelta(1)

    # -- application lifecycle ---------------------------------------------
    def start(self) -> "WorkbookHandler":
        if self.app is None:
            self.app = xw.App(visible=True, add_book=False)
        return self

    def __enter__(self) -> "WorkbookHandler":
        return self.start()

    def __exit__(self, exc_type, exc, tb) -> None:
        self.quit()

    def quit(self) -> None:
        if self.app is not None:
            try:
                self.app.quit()
            except Exception as e:  # noqa: BLE001 - best-effort teardown
                logger.warning("Error quitting Excel app: %s", e)
            finally:
                self.app = None

    # -- populate -----------------------------------------------------------
    def populate_all(self):
        """Populates the header cells and the employee table."""
        self.populate_row('A1', self._projectCode)
        self.populate_row('B1', self._dispoData['projname'])
        self.populate_row('A2', self._date.strftime("%B %d, %Y (%a)"))
        self.populate_row('R2', self._dispoData['inc'] / 100)
        self.populate_row('R3', self._dailyAVGData['avglen'])
        self.populate_row('R4', self._dispoData['mean'])
        self.populate_row('R1', self._dispoData['mean'])
        self.populate_row('U3', self._dailyAVGData['avgcph'])
        self.populate_row('U4', self._dailyAVGData['avgmph'])
        self.copy_rows()

    def populate_row(self, row: str, val):
        self._activeSheet.range(row).options(index=False, header=False).value = val

    def perf_swap(self):
        if self._projectCode.upper().endswith("C"):
            self.active_sheet = "C-PERF"
        else:
            self.active_sheet = "L-PERF"

    def populate_perf(self, sample_data, prel_data):
        # TODO Add INC and True Mean to data. It is in the lower table. True Mean = True Avg. Length and INC = True incidence
        # TODO Auto populate the date in the A column on the left side of the table using 16-JUL format
        try:
            self.active_sheet.range('A2').options(index=False, header=False).value = \
                f"{self._projectCode} {self._dispoData['projname'][0]}"
        except (ValueError, KeyError):
            logger.info("No %s", self.get_active_sheet_name())
            return

        columns = ['X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD']

        row = 30
        column = 'V'
        cell = None
        while True:
            cell = self.active_sheet.range(f'{column}{row}')
            if cell.value is None:
                break
            row += 3

        if cell is None:
            raise WorkbookError(
                "Cell is None. Please check the V column date cell start at 30 and increments by 3."
            )

        cell.value = self._date
        cell.number_format = "d-mmm"

        for key, value in prel_data['total'].items():
            column_index = 0 if key == '<>' else int(key) + 1
            self.active_sheet.range(f'{columns[column_index]}{row}').value = value

        for key, value in prel_data['co'].items():
            column_index = 0 if key == '<>' else int(key) + 1
            self.active_sheet.range(f'{columns[column_index]}{row + 1}').value = value

        row = 54
        column = 'V'
        cell = None
        while True:
            cell = self.active_sheet.range(f'{column}{row}')
            if cell.value is None:
                break
            row += 6

        if cell is None:
            raise WorkbookError(
                "Cell is None. Please check the V column date cell start at 54 and increments by 6."
            )

        cell.value = self._date
        cell.number_format = "d-mmm"

        for value in sample_data['used_sample_call_count']:
            self.active_sheet.range(f'{columns[value - 1]}{row}').value = \
                sample_data['used_sample_call_count'][value]

        for value in sample_data['live_connects_call_count']:
            self.active_sheet.range(f'{columns[value - 1]}{row + 1}').value = \
                sample_data['live_connects_call_count'][value]

        for value in sample_data['co_case_count']:
            self.active_sheet.range(f'{columns[value - 1]}{row + 3}').value = \
                sample_data['co_case_count'][value]

    def copy_rows(self) -> None:
        """Fills the employee data table A8:V{n}, pulling formula columns from row 8."""
        count = self._productionReportData.shape[0] + 7

        self.populate_row(f'A8:A{count}', self._productionReportData['eid'])
        self.populate_row(f'B8:B{count}', self._productionReportData['refname'])
        self.populate_row(f'C8:C{count}', self._productionReportData['recloc'])
        self.populate_row(f'D8:D{count}', self._productionReportData['tenure'])
        self.populate_row(f'E8:E{count}', self._productionReportData['hrs'])
        self.populate_row(f'F8:F{count}', self._productionReportData['connecttime'])

        self.populate_row(f'G8:G{count}', self._activeSheet.range("G8").formula)

        self.populate_row(f'H8:H{count}', self._productionReportData['pausetime'])

        self.populate_row(f'I8:I{count}', self._activeSheet.range("I8").formula)
        self.populate_row(f'J8:J{count}', self._activeSheet.range("J8").formula)
        self.populate_row(f'K8:K{count}', self._activeSheet.range("K8").formula)

        self.populate_row(f'L8:L{count}', self._productionReportData['cms'])

        self.populate_row(f'M8:M{count}', self._activeSheet.range("M8").formula)
        self.populate_row(f'N8:N{count}', self._activeSheet.range("N8").formula)
        self.populate_row(f'O8:O{count}', self._activeSheet.range("O8").formula)
        self.populate_row(f'P8:P{count}', self._activeSheet.range("P8").formula)

        self.populate_row(f'Q8:Q{count}', self._productionReportData['intal'])

        self.populate_row(f'R8:R{count}', self._activeSheet.range("R8").formula)

        self.populate_row(f'S8:S{count}', self._productionReportData['mph'])
        self.populate_row(f'T8:T{count}', self._productionReportData['totaldials'])

        self.populate_row(f'U8:U{count}', self._activeSheet.range("U8").formula)
        self.populate_row(f'V8:V{count}', self._activeSheet.range("V8").formula)

    def populate_expected_loi(self):
        """Pulls Expected LOI and populates the production report cell."""
        self._activeSheet.range('R1').options(index=False, header=False).value = self._dispoData['mean']

    def copy_sheet(self) -> None:
        """Runs the VBA macro that copies the template sheet (preserves formulas/styling)."""
        copySheet = self._wb.macro("Module1.copySheetTEST")  # (sheet: int)
        if self._projectCode[-1].upper() == "C":
            copySheet(2)
        else:
            copySheet(1)

    # -- filesystem / template ---------------------------------------------
    def check_path(self):
        production_dir = Path(os.environ['src']) / self._projectid / 'PRODUCTION'
        blank = Path(os.environ['src']) / 'PRODUCTION' / 'BLANK_Production.xlsm'
        if not production_dir.exists():
            production_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy(blank, self.path)
        elif not Path(self.path).exists():
            shutil.copy(blank, self.path)

    def create_sheet_name(self):
        return f"{self.project_code} {self._date.strftime('%m%d')}"

    # -- properties ---------------------------------------------------------
    @property
    def sheet_index(self):
        return self._activeSheet.index

    @property
    def project_id(self):
        return self._projectid

    @project_id.setter
    def project_id(self, projectid: Union[int, str]) -> None:
        self._projectid = str(projectid)

    @property
    def project_code(self):
        return self._projectCode

    @project_code.setter
    def project_code(self, projectCode) -> None:
        self._projectCode = str(projectCode)

    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, path: Path = None) -> None:
        if path is None:
            path = Path(os.environ['src']) / self._projectid / 'PRODUCTION' / \
                f"{self._projectid}_Production_Report.xlsm"
        self._path = path

    @property
    def workbook(self):
        return self._wb

    def set_workbook(self, path: str = None) -> None:
        """Opens the workbook. Raises WorkbookLockedError if it is locked/open elsewhere."""
        if path is None:
            path = self.path
        self._path = path
        try:
            self._wb = self.app.books.open(self._path)
        except Exception as e:  # noqa: BLE001 - xlwings/pywin32 raises generic errors
            msg = str(e).lower()
            if "being used" in msg or "locked" in msg or "permission" in msg or "in use" in msg:
                raise WorkbookLockedError(f"{self._path} is locked/open: {e}") from e
            raise WorkbookError(f"Could not open {self._path}: {e}") from e

    @property
    def active_sheet(self):
        return self._activeSheet

    @active_sheet.setter
    def active_sheet(self, activeSheet: Union[int, str] = None) -> None:
        if activeSheet is None:
            activeSheet = f"{self.get_active_sheet_name()}"
        self._activeSheet = self.workbook.sheets[activeSheet]
        self._wb.active = self._activeSheet

    @property
    def active_sheet_name(self):
        return self.active_sheet.name

    def get_active_sheet_name(self):
        if self.active_sheet.name is None:
            return 'No active sheet Name'
        return self.active_sheet.name

    def set_active_sheet_name(self, sheetName: str = None) -> None:
        """Renames the active sheet to ``{code} {MMDD}``.

        If that name already exists (a re-run of an already-generated project/date),
        raises :class:`SheetAlreadyExistsError` for the runner to log and skip.
        """
        if sheetName is None:
            sheetName = self._projectCode
        self._activeSheetName = f"{sheetName} {self._date.strftime('%m%d')}"
        try:
            self._wb.active.name = self._activeSheetName
        except Exception as e:  # noqa: BLE001 - xlwings raises generic COM errors
            raise SheetAlreadyExistsError(
                f"Sheet '{self._activeSheetName}' already exists (already generated?)"
            ) from e

    @property
    def data(self):
        return self._data

    @data.setter
    def data(self, data: tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]) -> None:
        productionReportData = data[0]
        productionReportData['intal'] = np.where(
            productionReportData['cms'] > 0, productionReportData['intal'], np.nan
        )
        productionReportData['mph'] = productionReportData['mph'].where(
            productionReportData['mph'] != 0.00, np.nan
        )
        self._productionReportData = productionReportData
        self._dispoData = data[1]
        self._dailyAVGData = data[2]
        self._data = data

    @property
    def production_data(self):
        return self._productionReportData

    @property
    def dispo_data(self):
        return self._dispoData

    @property
    def daily_avg_data(self):
        return self._dailyAVGData

    @property
    def date(self):
        return self._date

    @date.setter
    def date(self, date_):
        self._date = date_

    # -- save / close -------------------------------------------------------
    def save(self) -> None:
        self._wb.save(f"{self._path}")

    def close(self) -> None:
        if self._wb is not None:
            self._wb.close()
            self._wb = None
