import os
import shutil
from datetime import date, timedelta
from pathlib import Path
from typing import Union

import numpy as np
import pandas as pd
import xlwings as xw


class WorkbookHandler():
    app = xw.App(visible=True)

    def __init__(self):
        self._projectid = None
        self._projectCode = None
        self._path = None
        self._wb = None
        self._activeSheet = None
        self._activeSheetName = None
        self._productionData = pd.DataFrame
        self._date = date.today() - timedelta(1)

    def populate_all(self):
        """
        Populates all fields in sheet
        :return: None
        """
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
            self.set_active_sheet("CPERF")
        else:
            self.set_active_sheet("LPERF")

    def populate_perf(self):
        self._activeSheet.range('A2').options(index=False, header=False).value = self._dispoData['projname']

    def cperf(self):
        ...

    def lperf(self):
        ...

    def copy_rows(self) -> None:
        """
        Extends rows
        :param count: Number of rows to copy
        :return: None
        """

        count = self._productionReportData.shape[0] + 7

        self.populate_row(f'A8:A{count}', self._productionReportData['eid'])
        self.populate_row(f'B8:B{count}', self._productionReportData['refname'])
        self.populate_row(f'C8:C{count}', self._productionReportData['recloc'])
        self.populate_row(f'D8:D{count}', self._productionReportData['tenure'])
        self.populate_row(f'E8:E{count}', self._productionReportData['hrs'])
        self.populate_row(f'F8:F{count}', self._productionReportData['connecttime'])

        formula = self._activeSheet.range("G8").formula
        self.populate_row(f'G8:G{count}', formula)

        self.populate_row(f'H8:H{count}', self._productionReportData['pausetime'])

        formula = self._activeSheet.range("I8").formula
        self.populate_row(f'I8:I{count}', formula)

        formula = self._activeSheet.range("J8").formula
        self.populate_row(f'J8:J{count}', formula)

        formula = self._activeSheet.range("K8").formula
        self.populate_row(f'K8:K{count}', formula)

        self.populate_row(f'L8:L{count}', self._productionReportData['cms'])

        formula = self._activeSheet.range("M8").formula
        self.populate_row(f'M8:M{count}', formula)

        formula = self._activeSheet.range("N8").formula
        self.populate_row(f'N8:N{count}', formula)

        formula = self._activeSheet.range("O8").formula
        self.populate_row(f'O8:O{count}', formula)

        formula = self._activeSheet.range("P8").formula
        self.populate_row(f'P8:P{count}', formula)

        self.populate_row(f'Q8:Q{count}', self._productionReportData['intal'])

        formula = self._activeSheet.range("R8").formula
        self.populate_row(f'R8:R{count}', formula)

        self.populate_row(f'S8:S{count}', self._productionReportData['mph'])

        self.populate_row(f'T8:T{count}', self._productionReportData['totaldials'])

        formula = self._activeSheet.range("U8").formula
        self.populate_row(f'U8:U{count}', formula)

        formula = self._activeSheet.range("V8").formula
        self.populate_row(f'V8:V{count}', formula)

    def populate_expected_loi(self):
        """
        Pulls Expected LOI from <YEAR>PLANNER and populates the production report cell
        :return: None
        """
        self._activeSheet.range('R1').options(index=False, header=False).value = self._dispoData['mean']

    def copy_sheet(self) -> None:
        """
        Copies sheet
        :return: None
        """
        copySheet = self._wb.macro("Module1.copySheetTEST")  # Parameters (blankPath: str, path: str, projectid: str, sheet: int)

        if self._projectCode[-1].upper() == "C":
            copySheet(2)
        else:
            copySheet(1)
        del copySheet

    def check_path(self):
        if not os.path.exists(Path(f"{os.environ['src']}{self._projectid}/PRODUCTION/")):
            src = Path(f"{os.environ['src']}PRODUCTION/BLANK_Production.xlsm")
            os.mkdir(Path(f"{os.environ['src']}{self._projectid}/PRODUCTION/"))
            dst = self.get_path()
            shutil.copy(src, dst)
            del src, dst
        elif not os.path.exists(self.get_path()):
            src = Path(f"{os.environ['src']}PRODUCTION/BLANK_Production.xlsm")
            dst = self.get_path()
            shutil.copy(src, dst)
            del src, dst
        else:
            return

    def create_sheet_name(self):
        """
        Creates sheet Name
        :return: None
        """
        return f"{self.get_project_code()} {self._date.strftime('%m%d')}"

    def set_app(self):
        self.app = xw.App(visible=True)

    def set_project_id(self, projectid: Union[int, str]) -> None:
        """
        Sets project ID
        :param projectid: Int or String of Project ID
        :return: None
        """
        if type(projectid) == int:
            self._projectid = str(projectid)
        else:
            self._projectid = projectid

    def set_project_code(self, projectCode) -> None:
        """
        Sets Project Code

        Project code is defined as the projects directory ID.
        This is needed to identify projects such as #####C and name them accordingly.
        :param projectCode: Int or String of Project Code
        :return: None
        """
        if type(projectCode) == int:
            self._projectCode = str(projectCode)
        else:
            self._projectCode = projectCode

    def set_path(self, path: Path = None) -> None:
        """
        Sets path
        :param path: Filepath
        :return: None
        """
        if path is None:
            path = Path(f"{os.environ['src']}{self._projectid}/PRODUCTION/{self._projectid}_Production_Report.xlsm")
        self._path = path

    def set_workbook(self, path: str = None) -> None:
        """
        Sets workbook
        :param path: Filepath
        :return: None
        """
        if path is None:
            path = self.get_path()

        self._path = path
        self._wb = self.app.books.open(self._path)

    def set_active_sheet(self, activeSheet: Union[int, str] = None) -> None:
        """
        Sets active sheet
        :param activeSheet: Sheet Name
        :return: None
        """
        if activeSheet is None:
            activeSheet = f"{self.get_active_sheet_name()}"
        self._activeSheet = self.get_workbook().sheets[activeSheet]
        self._wb.active = self._activeSheet

    def set_active_sheet_name(self, sheetName: str = None) -> None:
        """
        Sets active sheet Name
        :param sheetName: Sheet Name
        :return: sheetName String
        """
        if sheetName is None:
            sheetName = self._projectCode
        self._activeSheetName = f"{sheetName} {self._date.strftime('%m%d')}"
        try:
            self._wb.active.name = self._activeSheetName
        except Exception as e:
            print("sheet name already taken", self._activeSheetName)
            raise (e, "sheet name already taken", self._activeSheetName)

    def set_data(self, data: tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]) -> None:
        """
        Sets Production Data
        :param data: Pandas Dataframe
        :return: None
        """
        productionReportData = data[0]
        productionReportData['intal'] = np.where(productionReportData['cms'] > 0, productionReportData['intal'], np.nan)
        productionReportData['mph'] = productionReportData['mph'].where(productionReportData['mph'] != 0.00, np.nan)
        dispoData = data[1]
        dailyAVGData = data[2]

        self._productionReportData = productionReportData
        self._dispoData = dispoData
        self._dailyAVGData = dailyAVGData
        self._data = data

    def set_date(self, date_):
        self._date = date_

    def get_workbook(self):
        """
        Gets workbook
        :return: Active Workbook
        """
        if self._wb is None:
            return 'No active workbook'
        return self._wb

    def get_active_sheet(self):
        """
        Gets active sheet
        :return: Active Sheet
        """
        if self._activeSheet is None:
            return 'No active sheet'
        return self._activeSheet

    def get_active_sheet_name(self):
        """
        Gets active sheet Name
        :return: Active Sheet Name
        """
        if self.get_active_sheet().name is None:
            return 'No active sheet Name'
        return self.get_active_sheet().name

    def get_project_id(self) -> str:
        """
        Gets project ID
        :return: Project ID
        """
        if self._projectid is None:
            return 'No project ID'
        return self._projectid

    def get_project_code(self) -> str:
        """
        Gets project Code
        :return: Project Code
        """
        if self._projectCode is None:
            return 'No project Code'
        return self._projectCode

    def get_path(self) -> str:
        """
        Gets path
        :return: Path
        """
        if self._path is None:
            return 'No path'
        return self._path

    def get_production_data(self) -> pd.DataFrame:
        """
        Gets Production Data
        :return: Pandas Dataframe
        """
        return self._productionReportData

    def get_dispo_data(self) -> pd.DataFrame:
        """
        Gets Dispo Data
        :return: Pandas Dataframe
        """
        return self._dispoData

    def get_daily_avg_data(self) -> pd.DataFrame:
        """
        Gets Daily Avg Data
        :return: Pandas Dataframe
        """
        return self._dailyAVGData

    def get_data(self) -> pd.DataFrame:
        """
        Gets Production Data
        :return: Pandas Dataframe
        """
        return self._data

    def get_date(self):
        """
        Gets date
        :return: date
        """
        return self._date

    def save(self) -> None:
        """
        Saves workbook
        :return: None
        """
        self._wb.save(f"{self._path}")

    def close(self) -> None:
        self._wb.close()
        # self.app.close()

    def app_quit(self) -> None:
        self.app.quit()
