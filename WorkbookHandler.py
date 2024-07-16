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
        self.populate_projectID()
        self.populate_project_name()  # this needs to be changed to E.g. (HEPA National (CRC))
        self.populate_date()
        self.populate_daily_inc()  # get from sql qry

        self.populate_avg_daily_loi()  # get from sql qry
        self.populate_avg_overall_loi()  # get from sql qry
        self.populate_expected_loi()  # get from planner
        self.populate_avg_cph()
        self.populate_avg_mph()
        self.copy_rows()

    def populate_projectID(self) -> None:
        """
        Populates Project ID
        :param project: Project ID
        :return: None
        """
        self._activeSheet.range('A1').options(index=False, header=False).value = self._projectCode

    def populate_project_name(self) -> None:
        """
        Populates project name
        :param projectName: Project name
        :return: None
        """
        # TODO run qry to pull project name
        self._activeSheet.range('B1').options(index=False, header=False).value = self._dispoData['projname']

    def populate_date(self) -> None:
        """
        Populates date
        :return: None
        """
        self._activeSheet.range('A2').options(index=False, header=False).value = self._date.strftime(
            "%B %d, %Y (%a)")

    def populate_daily_inc(self) -> None:
        """
        Populates Daily Incidence
        :param inc: Daily incedence
        :return: None
        """
        self._activeSheet.range('R2').options(index=False, header=False).value = self._dispoData['inc'] / 100

    def populate_avg_daily_loi(self) -> None:
        """
        Populates Avg. Daily LOI
        :param loi: AVG Daily LOI
        :return: None
        """
        self._activeSheet.range('R3').options(index=False, header=False).value = self._dailyAVGData['avglen']

    def populate_avg_overall_loi(self) -> None:
        """
        Populates Avg. Overall LOI
        :param loi: AVG Overall LOI
        :return: None
        """
        self._activeSheet.range('R4').options(index=False, header=False).value = self._dispoData['mean']

    def populate_avg_cph(self):
        """
        Populates Avg. CPH
        :return: None
        """
        self._activeSheet.range('U3').options(index=False, header=False).value = self._dailyAVGData['avgcph']

    def populate_avg_mph(self):
        """
        Populates Avg. MPH
        :return: None
        """
        self._activeSheet.range('U4').options(index=False, header=False).value = self._dailyAVGData['avgmph']

    def copy_rows(self) -> None:
        """
        Extends rows
        :param count: Number of rows to copy
        :return: None
        """
        count = self._productionReportData.shape[0] + 7
        self._activeSheet.range(f'A8:A{count}') \
            .options(index=False, header=False).value = self._productionReportData['eid']

        self._activeSheet.range(f'B8:B{count}') \
            .options(index=False, header=False).value = self._productionReportData['refname']

        self._activeSheet.range(f'C8:C{count}') \
            .options(index=False, header=False).value = self._productionReportData['recloc']

        self._activeSheet.range(f'D8:D{count}') \
            .options(index=False, header=False).value = self._productionReportData['tenure']

        self._activeSheet.range(f'E8:E{count}') \
            .options(index=False, header=False).value = self._productionReportData['hrs']

        self._activeSheet.range(f'F8:F{count}') \
            .options(index=False, header=False).value = self._productionReportData['connecttime']

        self._activeSheet.range(f'H8:H{count}') \
            .options(index=False, header=False).value = self._productionReportData['pausetime']

        self._activeSheet.range(f'L8:L{count}') \
            .options(index=False, header=False).value = self._productionReportData['cms']

        self._activeSheet.range(f'Q8:Q{count}') \
            .options(index=False, header=False).value = self._productionReportData['intal']

        self._activeSheet.range(f'S8:S{count}') \
            .options(index=False, header=False).value = self._productionReportData['mph']

        self._activeSheet.range(f'T8:T{count}') \
            .options(index=False, header=False).value = self._productionReportData['totaldials']

        formula = self._activeSheet.range("G8").formula
        self._activeSheet.range(f"G8:G{count}").formula = formula

        formula = self._activeSheet.range("I8").formula
        self._activeSheet.range(f"I8:I{count}").formula = formula

        formula = self._activeSheet.range("J8").formula
        self._activeSheet.range(f"J8:J{count}").formula = formula

        formula = self._activeSheet.range("K8").formula
        self._activeSheet.range(f"K8:K{count}").formula = formula

        formula = self._activeSheet.range("M8").formula
        self._activeSheet.range(f"M8:M{count}").formula = formula

        formula = self._activeSheet.range("N8").formula
        self._activeSheet.range(f"N8:N{count}").formula = formula

        formula = self._activeSheet.range("O8").formula
        self._activeSheet.range(f"O8:O{count}").formula = formula

        formula = self._activeSheet.range("P8").formula
        self._activeSheet.range(f"P8:P{count}").formula = formula

        formula = self._activeSheet.range("R8").formula
        self._activeSheet.range(f"R8:R{count}").formula = formula

        formula = self._activeSheet.range("U8").formula
        self._activeSheet.range(f"U8:U{count}").formula = formula

        formula = self._activeSheet.range("V8").formula
        self._activeSheet.range(f"V8:V{count}").formula = formula

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
