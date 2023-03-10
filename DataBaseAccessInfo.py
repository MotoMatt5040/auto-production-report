from configparser import ConfigParser
import os


class DataBaseAccessInfo:

    # region constants for server info
    config_object = ConfigParser()
    config_path = f"C:/Users/{os.getlogin()}/AppData/Local/AutoProductionReport/config.ini"
    config_object.read(config_path)


    info = config_object["DATABASE ACCESS INFO"]

    CORESERVER = info["coreserver"]
    CC3SERVER = info["cc3server"]
    ARCHIVECC3SERVER = info["archivecc3server"]
    TFSERVER = info["tfserver"]

    COREUSER = info["coreuser"]
    CC3USER = info["cc3user"]
    ARCHIVECC3USER = info["archivecc3user"]
    TFUSER = info["tfuser"]

    COREPASSWORD = info["corepassword"]
    CC3PASSWORD = info["cc3password"]
    ARCHIVECC3PASSWORD = info["archivecc3password"]
    TFPASSWORD = info["tfpassword"]

    CALIGULA = info["caligula"]
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
            print(err)

    def get_info(self):
        URI = '*****\nDriver: {}\nServer Address: {}\nDatabase Name: {}\nUserID: {}\nPassword: {}\n*****'.format(
            self._driver, self._servername, self._database, self._userid, self._password
        )
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
