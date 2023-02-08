import pandas as pd
import os
import win32com.client
import numpy as np

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
        self._wb = None
        self._activeSheet = None
        self._activeSheetTitle = None
        self._projectid = None
        self._path = None

    def populateProjectNumber(self, project: Union[str, int]):
        """
        Populates project number
        :param project: Project number
        :return: None
        """
        self._activeSheet.cell(row=1, column=1).value = project

    def populateProjectName(self, projectName: str):
        """
        Populates project name
        :param projectName: Project name
        :return: None
        """
        self._activeSheet.cell(row=1, column=2).value = projectName

    def populateDate(self):
        """
        Populates date
        :return: None
        """
        self._activeSheet.cell(row=2, column=1).value = datetime.now().strftime("%B %d, %Y (%a)")

    def populateDailyINC(self, inc: float):
        """
        Populates Daily Incidence
        :param inc: Daily incedence
        :return: None
        """
        self._activeSheet.cell(row=2, column=18).value = inc

    def populateAVGDailyLOI(self, loi):
        """
        Populates Avg. Daily LOI
        :param loi: AVG Daily LOI
        :return: None
        """
        self._activeSheet.cell(row=3, column=18).value = loi

    def populateAVGOverallLOI(self, loi):
        """
        Populates Avg. Overall LOI
        :param loi: AVG Overall LOI
        :return: None
        """
        self._activeSheet.cell(row=4, column=18).value = loi

    def copyRows(self, count):
        """
        Extends rows
        :param count: Number of rows to copy
        :return: None
        """
        for i in range(count-1):
            row = i + 9
            self._activeSheet[f"G{row}"].value = f"=(E{row} - F{row}) * 60"
            self._activeSheet[f"I{row}"].value = f"=(H{row}) * 60"
            self._activeSheet[f"J{row}"].value = f"=IFERROR(H{row}/E{row},0)"
            self._activeSheet[f"K{row}"].value = f"=IFERROR(IF(G{row}>0,(G{row}+I{row})/(E{row}*60),I{row}/(E{row}*60)),0)"
            self._activeSheet[f"M{row}"].value = f"=IFERROR(ROUND(E{row}*$U$1,2),0)"
            self._activeSheet[f"N{row}"].value = f"=L{row}-M{row}"
            self._activeSheet[f"O{row}"].value = f"=RANK(N{row},$N$8:$N$140)"
            self._activeSheet[f"P{row}"].value = f"=IFERROR(T{row}/$AC$85,0)"
            self._activeSheet[f"R{row}"].value = f"=IFERROR(L{row}/E{row},0)"
            self._activeSheet[f"V{row}"].value = f"=IFERROR(T{row}/E{row},0)"
        print(self._activeSheet.title, "extending rows")

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
        self._wb.app.macro(f"copy_LL")
        print(self._path)
        self.save()
        # self.save()
        # xl = win32com.client.Dispatch("Excel.Application")
        # wb = xl.workbooks.open(self._path)
        # xl.run(f"{self._projectid}_Production_ReportTEST.xlsm!Module1.copy_LL")
        # xl.visible = True
        # wb.Close(SaveChanges=1)
        # xl.Quit()
        # del xl, wb
        # self.setWorkbook(self._path)
        # if self._wb.active.title != f"{self._projectid} {date.today().strftime('%m%d')}":
        #     pass
        # test = self._wb.copy_worksheet(self._wb.active)
        # help(self._wb.copy_worksheet)

    def setPath(self, path: str) -> None:
        """
        Sets path
        :param path: Filepath
        :return: None
        """
        self._path = path

    def setWorkbook(self, path) -> None:
        """
        Sets workbook
        :param path: Filepath
        :return: None
        """
        self.app = xw.App(visible=True)
        self._path = path
        self._wb = self.app.books.open(self._path)
        # self._wb = load_workbook(path)

    def setActiveSheet(self, activeSheet: str) -> None:
        """
        Sets active sheet
        :param active: Sheet title
        :return: None
        """
        self._activeSheet = self._wb.sheets[activeSheet]
        self._wb.active = self._wb.sheets[activeSheet]

    def setActiveSheetTitle(self, sheetTitle: str):
        """
        Sets active sheet title
        :param sheetTitle: Sheet title
        :return: sheetTitle String
        """
        self._activeSheetTitle = f"{sheetTitle} {date.today().strftime('%m%d')}"
        self._wb.active.title = self._activeSheetTitle

    def setProjectID(self, projectid: Union[int, str]):
        """
        Sets project ID
        :param projectid: Int or String of Project ID
        :return: None
        """
        if type(projectid) == int:
            self._projectid = str(projectid)
        else:
            self._projectid = projectid

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
        if self._activeSheet is None:
            return 'No active sheet'
        return self._activeSheet

    def getActiveSheetTitle(self):
        """
        Gets active sheet title
        :return: Active Sheet Title
        """
        if self._activeSheetTitle is None:
            return 'No active sheet title'
        return self._activeSheetTitle

    def getProjectID(self) -> Union[int, str]:
        """
        Gets project ID
        :return: Project ID
        """
        if self._projectid is None:
            return 'No project ID'
        return self._projectid

    def save(self) -> None:
        """
        Saves workbook
        :return: None
        """
        # TODO Replace with path instead of title
        self._wb.save(f"{self._path}")

    def close(self) -> None:
        self._wb.close()
        self.app.kill()
