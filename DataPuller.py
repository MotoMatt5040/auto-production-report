import logging
from logging.handlers import RotatingFileHandler

import bcrypt
import numpy as np
import pandas
import pandas as pd
from sqlalchemy import create_engine
import math

import DataBaseAccessInfo as dbai
import SQLDictionary as sqld

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(thread)d:%(threadName)s:%(name)s:%(funcName)s:%(levelname)s:%(message)s')
file_handler = RotatingFileHandler('logs\DataPuller.log', mode='a', maxBytes=1024, backupCount=1)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.propagate = False

class DataPuller:

    #  initialize database access info to connect to database
    DB_INFO = dbai.DataBaseAccessInfo(
        driver='ODBC17',
        servername='COREServer',
        database='Caligula',
        userid='COREUser',
        password='COREPassword'
    )

    #  Build the connection string
    SQL_CONNECTION = 'DRIVER={};SERVER={};DATABASE={};UID={};PWD={}'.format(
        DB_INFO.get_driver(),
        DB_INFO.get_server(),
        DB_INFO.get_database(),
        DB_INFO.get_user_id(),
        DB_INFO.get_password()
    )
    dbcon = f'mssql+pyodbc:///?odbc_connect={SQL_CONNECTION}'
    engine = create_engine(dbcon)

    def pull_data(self, qry: str):
        try:
            cnxn = self.engine.connect()
            df = pd.read_sql_query(qry, cnxn)
            cnxn.close()
            return df
        except Exception as err:
            logger.critical(err)
        return []

