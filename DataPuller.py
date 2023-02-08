import logging
import pandas as pd
from sqlalchemy import create_engine
from logging.handlers import RotatingFileHandler

import DataBaseAccessInfo
import SQLDictionary

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(thread)d:%(threadName)s:%(name)s:%(funcName)s:%(levelname)s:%(message)s')
file_handler = RotatingFileHandler('logs\DataPuller.log', mode='a', maxBytes=1024, backupCount=1)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.propagate = False

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

    def activeProjectIDs(self):
        try:
            cnxn = self.engine.connect()
            df = pd.read_sql_query(self.sqld.project_id_list(), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            logger.critical(err)
        return []

    def gpcph(self):
        try:
            cnxn = self.engine.connect()
            df = pd.read_sql_query(self.sqld.gpcph(), cnxn)
            cnxn.close()
            del cnxn
            return df
        except Exception as err:
            logger.critical(err)
        return []
