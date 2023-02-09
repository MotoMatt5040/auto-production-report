import pandas as pd
import os
import shutil

from typing import Union
from datetime import date, datetime
import xlwings as xw
from configparser import ConfigParser


config_object = ConfigParser()
config_object.read('config.ini')
file_paths = config_object['FILE PATHS']
del config_object

class WorkbookHandler():


    def __init__(self):
        self._projectid = None
        self._projectCode = None
        self._path = None
        self._wb = None
        self._activeSheet = None
        self._activeSheetName = None

    def populateAll(self):
        """
        Populates all fields in sheet
        :return: None
        """
        self.populateProjectID()
        self.populateProjectName(self._projectCode)  # this needs to be changed to E.g. (HEPA National (CRC))
        self.populateDate()
        self.populateDailyINC(0)  # get from sql qry
        self.populateExpectedLOI()  # get from planner
        self.populateAVGDailyLOI(0)   # get from sql qry
        self.populateAVGOverallLOI(0)  # get from sql qry

    def populateProjectID(self) -> None:
        """
        Populates Project ID
        :param project: Project ID
        :return: None
        """
        self._activeSheet.cell(row=1, column=1).value = self._projectCode

    def populateProjectName(self, projectName: str) -> None:
        """
        Populates project name
        :param projectName: Project name
        :return: None
        """
        # TODO run qry to pull project name
        self._activeSheet.cell(row=1, column=2).value = projectName

    def populateDate(self) -> None:
        """
        Populates date
        :return: None
        """
        self._activeSheet.cell(row=2, column=1).value = datetime.now().strftime("%B %d, %Y (%a)")

    def populateDailyINC(self, inc: float) -> None:
        """
        Populates Daily Incidence
        :param inc: Daily incedence
        :return: None
        """
        self._activeSheet.cell(row=2, column=18).value = inc

    def populateAVGDailyLOI(self, loi) -> None:
        """
        Populates Avg. Daily LOI
        :param loi: AVG Daily LOI
        :return: None
        """
        self._activeSheet.cell(row=3, column=18).value = loi

    def populateAVGOverallLOI(self, loi) -> None:
        """
        Populates Avg. Overall LOI
        :param loi: AVG Overall LOI
        :return: None
        """
        self._activeSheet.cell(row=4, column=18).value = loi

    def copyRows(self, count) -> None:
        """
        Extends rows
        :param count: Number of rows to copy
        :return: None
        """
        count += 7  # First row is 8, so we must add 7 to be inclusive
        formula = self.getActiveSheet().range("G8").formula
        self.getActiveSheet().range(f"G8:G{count}").formula = formula

        formula = self.getActiveSheet().range("I8").formula
        self.getActiveSheet().range(f"I8:I{count}").formula = formula

        formula = self.getActiveSheet().range("J8").formula
        self.getActiveSheet().range(f"J8:J{count}").formula = formula

        formula = self.getActiveSheet().range("K8").formula
        self.getActiveSheet().range(f"K8:K{count}").formula = formula

        formula = self.getActiveSheet().range("M8").formula
        self.getActiveSheet().range(f"M8:M{count}").formula = formula

        formula = self.getActiveSheet().range("N8").formula
        self.getActiveSheet().range(f"N8:N{count}").formula = formula

        formula = self.getActiveSheet().range("O8").formula
        self.getActiveSheet().range(f"O8:O{count}").formula = formula

        formula = self.getActiveSheet().range("P8").formula
        self.getActiveSheet().range(f"P8:P{count}").formula = formula

        formula = self.getActiveSheet().range("R8").formula
        self.getActiveSheet().range(f"R8:R{count}").formula = formula

        formula = self.getActiveSheet().range("V8").formula
        self.getActiveSheet().range(f"V8:V{count}").formula = formula

    def populateExpectedLOI(self, projectid: Union[int, str] = '') -> str:
        """
        Pulls Expected LOI from 2020PLANNER and populates the production report cell
        :return: None
        """
        projectid = '12527C'
        if self._activeSheet['R1'].value is None:
            try:
                df = pd.read_excel(f"{file_paths['PLANNER']}{date.today().strftime('%Y')}PLANNER.xls")
                if type(projectid) == int:
                    df = df.query(f"`Unnamed: 1` == {projectid}")
                else:
                    df = df.query(f"`Unnamed: 1` == '{projectid}'")
                df = df.dropna(axis=1, how='all')
                expectedLOI = df.iloc[0]['Unnamed: 16']
            except Exception as err:
                print(err)
                self._activeSheet['B1'].value = 'ERROR IN READING PLANNER'

    def copySheet(self) -> None:
        """
        Copies sheet
        :return: None
        """
        copySheet = self._wb.macro(f"Module1.copySheet") # Parameters (blankPath: str, path: str, projectid: str, sheet: int)
        if self._projectCode[-1].upper() == "C":
            copySheet(
                f"{file_paths['SRC']}PRODUCTION/BLANK_Production.xlsm",
                f"{file_paths['SRC']}{self._projectid}/PRODUCTION/{self._projectid}_Production_ReportTEST.xlsm",
                self._projectid,
                2
            )
        else:
            copySheet(
                f"{file_paths['SRC']}PRODUCTION/BLANK_Production.xlsm",
                f"{file_paths['SRC']}{self._projectid}/PRODUCTION/{self._projectid}_Production_ReportTEST.xlsm",
                self._projectid,
                1
            )
        del copySheet

    def checkPath(self):
        if not os.path.exists(f"{file_paths['SRC']}{self._projectid}/PRODUCTION/"):
            src = f"{file_paths['SRC']}PRODUCTION/BLANK_Production.xlsm"
            os.mkdir(f"{file_paths['SRC']}{self._projectid}/PRODUCTION/")
            dst = self.getPath()
            shutil.copy(src, dst)
            del src, dst
        elif not os.path.exists(self.getPath()):
            src = f"{file_paths['SRC']}PRODUCTION/BLANK_Production.xlsm"
            dst = self.getPath()
            shutil.copy(src, dst)
            del src, dst
        else:
            return

    def createSheetName(self):
        """
        Creates sheet Name
        :return: None
        """
        return f"{self.getProjectCode()} {date.today().strftime('%m%d')}"

    def setProjectID(self, projectid: Union[int, str]) -> None:
        """
        Sets project ID
        :param projectid: Int or String of Project ID
        :return: None
        """
        if type(projectid) == int:
            self._projectid = str(projectid)
        else:
            self._projectid = projectid

    def setProjectCode(self, projectCode) -> None:
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

    def setPath(self, path: str = None) -> None:
        """
        Sets path
        :param path: Filepath
        :return: None
        """
        if path is None:
            path = f"{file_paths['SRC']}{self._projectid}/PRODUCTION/{self._projectid}_Production_ReportTEST.xlsm"
        self._path = path


    def setWorkbook(self, path: str = None) -> None:
        """
        Sets workbook
        :param path: Filepath
        :return: None
        """
        if path is None:
            path = self.getPath()
        self.app = xw.App(visible=True)
        self._path = path
        self._wb = self.app.books.open(self._path)
        # self._wb = load_workbook(path)

    def setActiveSheet(self, activeSheet: Union[int, str] = None) -> None:
        """
        Sets active sheet
        :param activeSheet: Sheet Name
        :return: None
        """
        # print(f"active Sheet: -{activeSheet}-")
        # if not activeSheet:
        #     self._activeSheet = self._wb.sheets[f"{}"]
        if activeSheet is None:
            activeSheet = f"{self.getActiveSheetName()}"
        self._wb.active = self.getWorkbook().sheets[activeSheet]

    def setActiveSheetName(self, sheetName: str = None) -> None:
        """
        Sets active sheet Name
        :param sheetName: Sheet Name
        :return: sheetName String
        """
        print(self._projectCode)
        if sheetName is None:
            sheetName = self._projectCode
        self._activeSheetName = f"{sheetName} {date.today().strftime('%m%d')}"
        self._wb.active.name = self._activeSheetName

    def getWorkbook(self):
        """
        Gets workbook
        :return: Active Workbook
        """
        if self._wb is None:
            return 'No active workbook'
        return self._wb

    def getActiveSheet(self):
        """
        Gets active sheet
        :return: Active Sheet
        """
        if self._wb.active is None:
            return 'No active sheet'
        return self._wb.active

    def getActiveSheetName(self):
        """
        Gets active sheet Name
        :return: Active Sheet Name
        """
        if self.getActiveSheet().name is None:
            return 'No active sheet Name'
        return self.getActiveSheet().name

    def getProjectID(self) -> str:
        """
        Gets project ID
        :return: Project ID
        """
        if self._projectid is None:
            return 'No project ID'
        return self._projectid

    def getProjectCode(self) -> str:
        """
        Gets project Code
        :return: Project Code
        """
        if self._projectCode is None:
            return 'No project Code'
        return self._projectCode

    def getPath(self) -> str:
        """
        Gets path
        :return: Path
        """
        if self._path is None:
            return 'No path'
        return self._path

    def save(self) -> None:
        """
        Saves workbook
        :return: None
        """
        # TODO Replace with path instead of Name
        self._wb.save(f"{self._path}")

    def close(self) -> None:
        self._wb.close()
        self.app.kill()
