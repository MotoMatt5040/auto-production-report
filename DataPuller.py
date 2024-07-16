import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy import text
import webbrowser
import pyodbc
from DataBaseAccessInfo import DataBaseAccessInfo
from SQLDictionary import SQLDictionary
import traceback


def error_log(err):
    print("\n\n\n")
    print("*" * 350)
    print(err)
    print("\n")
    print("*" * 50)
    print("\n")
    print(traceback.format_exc())
    # input("Press Enter to continue...")
    print("*" * 150)
    print("\n\n\n")

class DataPuller:

    def __init__(self):
        '''Data Puller class'''
        self.sqld = SQLDictionary()

        #  initialize database access info to connect to database
        self.dbai = DataBaseAccessInfo()

        #  Build the connection string
        self.SQL_CONNECTION = 'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
            self.dbai.driver,
            self.dbai.server,
            self.dbai.database,
            self.dbai.user_id,
            self.dbai.password
        )

        self.dbcon = f'mssql+pyodbc:///?odbc_connect={self.SQL_CONNECTION}'
        # connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": self.SQL_CONNECTION})
        try:
            # self.engine = create_engine(self.dbcon, pool_pre_ping=True)
            self.engine = create_engine(self.dbcon)
        except Exception as err:
            error_log(err)

    def __str__(self):
        return self.SQL_CONNECTION

    def active_project_ids(self):
        try:
            try:
                cnxn = self.engine.connect()
            except Exception:
                print("No ODBC Driver detected. Please install Microsoft ODBC Driver 17 for SQL Server from "
                      "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
                webbrowser.open(
                    "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
            df = pd.read_sql_query(text(self.sqld.project_id_list()), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            raise(err)
            print(err)
        return []

    def production_report(self, projectid, date) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        try:
            try:
                cnxn = self.engine.connect()
            except Exception:
                print("No ODBC Driver detected. Please install Microsoft ODBC Driver 17 for SQL Server from "
                      "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
                webbrowser.open(
                    "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
                input("Press Enter to continue...")
            data = self.sqld.production_report(projectid, date)
            productionReport = pd.read_sql_query(text(data[0]), cnxn)
            productionReportDispo = pd.read_sql_query(text(data[1]), cnxn)
            productionReportDailyAVG = pd.read_sql_query(text(data[2]), cnxn)
            cnxn.close()
            del cnxn
            return productionReport, productionReportDispo, productionReportDailyAVG
        except Exception as err:
            print(err)
            input("Press Enter to continue...")
        return []
