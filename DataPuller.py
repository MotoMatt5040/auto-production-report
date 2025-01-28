from typing import Dict

import pandas as pd
from sqlalchemy import text
from DataBaseAccessInfo import DataBaseAccessInfo
from SQLDictionary import SQLDictionary
import traceback
import os
from utils.logger_config import logger
from datetime import datetime

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
            cnxn = self.dbai.connect_engine()
            df = pd.read_sql_query(text(self.sqld.project_id_list()), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            raise err

    def production_report(self, projectid, date: datetime.date) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        try:
            cnxn = self.dbai.connect_engine()
            data = self.sqld.production_report(projectid, date)
            productionReport = pd.read_sql_query(text(data[0]), cnxn)
            productionReportDispo = pd.read_sql_query(text(data[1]), cnxn)
            productionReportDailyAVG = pd.read_sql_query(text(data[2]), cnxn)
            cnxn.close()
            del cnxn
            return productionReport, productionReportDispo, productionReportDailyAVG
        except Exception as err:
            raise err

    def get_voxco_project_database(self, project_number: str):
        try:
            self.dbai.find_voxco_project_database()
            cnxn = self.dbai.connect_engine()
            df = pd.read_sql_query(text(f"{os.environ['voxco_project_database']}{project_number}'"), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            raise err

    def get_voxco_data_sample(self, database: str, date):
        """
        Retrieves and processes sample data from a Voxco database.

        This method connects to the specified Voxco database, retrieves data for the given date,
        and processes the data into categorized DataFrames based on `HisCallNumber` values.
        It also filters these DataFrames to provide counts for:
        - All rows (`used_sample_call_count`).
        - Rows with valid `HisCaseResult` values (`live_connects_call_count`).
        - Rows with `HisCaseResult` equal to `'CO'` (`co_case_count`).

        Args:
            database (str): The name of the database to connect to.
            date (str): The date for which to retrieve data, formatted as 'YYYY-MM-DD'.

        Returns:
            dict: A dictionary containing three nested dictionaries:
                - `used_sample_call_count`: A dictionary where keys are `HisCallNumber` values (1-6)
                  and the values are row counts of the corresponding DataFrames.
                - `live_connects_call_count`: A dictionary where keys are `HisCallNumber` values (1-6)
                  and the values are row counts of the filtered DataFrames with valid `HisCaseResult` values.
                - `co_case_count`: A dictionary where keys are `HisCallNumber` values (1-6)
                  and the values are row counts of the filtered DataFrames with `HisCaseResult` equal to `'CO'`.

        Raises:
            Exception: If any error occurs during database connection, query execution, or data processing.

        Example:
            >>> data = get_voxco_data_sample("VoxcoDB", "2025-01-01")
            >>> print(data['used_sample_call_count'][1])  # Row count for HisCallNumber 1
            >>> print(data['live_connects_call_count'][1])  # Row count for HisCallNumber 1 with valid results
            >>> print(data['co_case_count'][1])  # Row count for HisCallNumber 1 with 'CO' result

        Notes:
            - `HisCallNumber` values 1-5 are processed into separate DataFrames, and values >= 6 are grouped
              into `used_sample_call_count_6`.
            - Valid `HisCaseResult` values include codes such as 'OK', 'CO', '05', '06', '07', '10', etc.
            - Logs the row counts for the main DataFrame and processed DataFrames for debugging purposes.
            - The method processes and filters the data locally for improved performance instead of
              executing multiple queries.
        """
        try:
            # self.dbai.voxco_db(database)
            cnxn = self.dbai.connect_engine()
            df = pd.read_sql_query(text(self.sqld.voxco_sample_data(database, date)), cnxn)
            cnxn.close()
            del cnxn

            # NOTE: This is faster than running queries. It takes double the time to run multiple queries.

            filters = [1, 2, 3, 4, 5]
            dfs = {f"used_sample_call_count_{n}": df[df['HisCallNumber'] == n] for n in filters}
            dfs['used_sample_call_count_6'] = df[df['HisCallNumber'] >= 6]

            valid_case_results = [
                'OK', 'CO', '05', '06', '07', '10', '12', '13', '14', '16', '22', '23', '25',
                '26', '80', '90', '91', '92', '93', '94', '95', '96', '97', '98', '99'
            ]

            dfls = {
                key.replace('used_sample_call_count_', 'live_connects_call_count_'): value[value['HisCaseResult'].isin(valid_case_results)]
                for key, value in dfs.items()
            }

            dfls_CO = {
                key.replace('used_sample_call_count_', 'co_case_count_'): value[value['HisCaseResult'] == 'CO']
                for key, value in dfs.items()
            }

            logger.debug(df.shape[0])
            logger.debug(dfs['used_sample_call_count_1'].shape[0])
            logger.debug(dfs['used_sample_call_count_2'].shape[0])
            logger.debug(dfs['used_sample_call_count_3'].shape[0])
            logger.debug(dfls['live_connects_call_count_1'].shape[0])
            logger.debug(dfls['live_connects_call_count_2'].shape[0])
            logger.debug(dfls['live_connects_call_count_3'].shape[0])

            data = {
                'used_sample_call_count': {n: dfs[f'used_sample_call_count_{n}'].shape[0] for n in range(1, 7)},
                'live_connects_call_count': {n: dfls[f'live_connects_call_count_{n}'].shape[0] for n in range(1, 7)},
                'co_case_count': {n: dfls_CO[f'co_case_count_{n}'].shape[0] for n in range(1, 7)}
            }

            return data
        except Exception as err:
            raise err

    def get_prel_data_sample(self, database: str, date):
        try:
            # self.dbai.voxco_db(database)
            cnxn = self.dbai.connect_engine()
            logger.debug(cnxn)
            logger.debug(self.dbai.get_info())
            df = pd.read_sql_query(text(self.sqld.voxco_prel_sample_data(database, date)), cnxn)
            cnxn.close()
            del cnxn

            # NOTE: This is faster than running queries. It takes double the time to run multiple queries.

            filters = [0, 1, 2, 3, 4, 5]
            dfs = {'prel_<>': df[df['RpsContent'].isna()]}
            dfs.update({f"prel_{n}": df[df['RpsContent'] == str(n)] for n in filters})

            dfco = {
                key.replace('prel_', 'co_'): value[value['HisCaseResult'] == 'CO']
                for key, value in dfs.items()
            }

            data = {
                'total': {key.replace('prel_', ''): dfs[key].shape[0] for key in dfs},
                'co': {key.replace('co_', ''): dfco[key].shape[0] for key in dfco}
            }

            logger.debug(data)

            return data
        except Exception as err:
            raise err





