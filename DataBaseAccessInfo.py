import logging
from logging.handlers import RotatingFileHandler

from configparser import ConfigParser

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(name)s:%(funcName)s:%(levelname)s:%(message)s')
file_handler = RotatingFileHandler('logs\DataBaseAccessInfo.log', mode='a', maxBytes=1024, backupCount=1)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.propagate = False


class DataBaseAccessInfo:
    # region constants for server info
    config_object = ConfigParser()
    config_object.read("config.ini")

    info = config_object['DATABASE ACCESS INFO']

    CORESERVER = info["CORESERVER"]
    CC3SERVER = info["CC3SERVER"]
    ARCHIVECC3SERVER = info["ARCHIVECC3SERVER"]
    TFSERVER = info["TFSERVER"]

    COREUSER = info["COREUSER"]
    CC3USER = info["CC3USER"]
    ARCHIVECC3USER = info["ARCHIVECC3USER"]
    TFUSER = info["TFUSER"]

    COREPASSWORD = info["COREPASSWORD"]
    CC3PASSWORD = info["CC3PASSWORD"]
    ARCHIVECC3PASSWORD = info["ARCHIVECC3PASSWORD"]
    TFPASSWORD = info["TFPASSWORD"]

    CALIGULA = info["CALIGULA"]
    del config_object

    SERVERNAME = {
        'COREServer': CORESERVER,
        'CC3Server': CC3SERVER,
        'ArchiveCC3Server': ARCHIVECC3SERVER,
        'TFServer': TFSERVER
    }

    DATABASE = {
        'Caligula': CALIGULA
    }

    USERID = {
        'COREUser': COREUSER,
        'CC3User': CC3USER,
        'ArchiveCC3User': ARCHIVECC3USER,
        'TFUser': TFUSER
    }

    PASSWORD = {
        'COREPassword': COREPASSWORD,
        'CC3Password': CC3PASSWORD,
        'ArchiveCC3Password': ARCHIVECC3PASSWORD,
        'TFPassword': TFPASSWORD
    }

    DRIVER = {
        'ODBC17': '{ODBC Driver 17 for SQL Server}'
    }
    # endregion constants for server info

    def __init__(self, driver, servername: str, database: str, userid: str, password: str):
        try:
            self._driver = self.DRIVER[driver]
            self._servername = self.SERVERNAME[servername]
            self._database = self.DATABASE[database]
            self._userid = self.USERID[userid]
            self._password = self.PASSWORD[password]
        except Exception as err:
            logger.critical(err)

    def get_info(self):
        URI = '*****\nDriver: {}\nServer Address: {}\nDatabase Name: {}\nUserID: {}\nPassword: {}\n*****'.format(
            self._driver, self._servername, self._database, self._userid, self._password
        )
        logger.info(URI)
        return URI

    def get_driver(self):
        return self._driver

    def get_server(self):
        return self._servername

    def get_database(self):
        return self._database

    def get_user_id(self):
        return self._userid

    def get_password(self):
        return self._password
