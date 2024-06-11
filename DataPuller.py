import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy import text
import webbrowser
import pyodbc
import DataBaseAccessInfo
import SQLDictionary

class DataPuller:
    sqld = SQLDictionary.SQLDictionary()

    #  initialize database access info to connect to database
    dbai = DataBaseAccessInfo.DataBaseAccessInfo(
        driver='ODBC17',
        servername='COREServer',
        database='Caligula',
        userid='COREUser',
        password='COREPassword'
    )

    #  Build the connection string
    SQL_CONNECTION = 'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
        dbai.get_driver(),
        dbai.get_server(),
        dbai.get_database(),
        dbai.get_user_id(),
        dbai.get_password()
    )
    del dbai
    dbcon = f'mssql+pyodbc:///?odbc_connect={SQL_CONNECTION}'
    engine = create_engine(dbcon)
    del dbcon

    def active_project_ids(self):
        try:
            try:
                cnxn = self.engine.connect()
            except Exception:
                print("No ODBC Driver detected. Please install Microsoft ODBC Driver 17 for SQL Server from "
                      "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
                webbrowser.open(
                    "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
                input("Press Enter to continue...")
            df = pd.read_sql_query(text(self.sqld.project_id_list()), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            input("Press Enter to continue...")
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
