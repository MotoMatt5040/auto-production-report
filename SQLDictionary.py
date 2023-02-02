import logging
from datetime import date, datetime, timedelta
from logging.handlers import RotatingFileHandler

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(name)s:%(funcName)s:%(levelname)s:%(message)s')
file_handler = RotatingFileHandler('logs\SQLDictionary.log', mode='a', maxBytes=1024)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.propagate = False

class SQLDictionary:



    def project_id_list(self, qry):

        projectids = f"{qry}'{str(date.today() - timedelta(1))}'"
        return projectids