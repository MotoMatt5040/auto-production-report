import pandas as pd
from sqlalchemy import text
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


    def active_project_ids(self):
        try:
            cnxn = self.dbai.connection
            df = pd.read_sql_query(text(self.sqld.project_id_list()), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            raise err

    def production_report(self, projectid, date) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        try:
            cnxn = self.dbai.connection
            data = self.sqld.production_report(projectid, date)
            productionReport = pd.read_sql_query(text(data[0]), cnxn)
            productionReportDispo = pd.read_sql_query(text(data[1]), cnxn)
            productionReportDailyAVG = pd.read_sql_query(text(data[2]), cnxn)
            cnxn.close()
            del cnxn
            return productionReport, productionReportDispo, productionReportDailyAVG
        except Exception as err:
            raise err




