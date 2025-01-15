import os

from sqlalchemy import create_engine
import webbrowser


class DataBaseAccessInfo:
    _isInstance = None

    def __new__(cls):
        if cls._isInstance is None:
            cls._isInstance = super(DataBaseAccessInfo, cls).__new__(cls)
            cls._initialize(cls._isInstance)

        return cls._isInstance

    @staticmethod
    def _initialize(instance):
        try:
            # Initialize the instance if it hasn't been initialized
            if not hasattr(instance, '_initialized'):
                instance._driver = '{ODBC Driver 17 for SQL Server}'
                instance._servername = os.environ['coreserver']
                instance._database = os.environ['caligula']
                instance._userid = os.environ['coreuser']
                instance._password = os.environ['corepassword']
                instance._initialized = True
        except Exception as err:
            print(f"Initialization error: {err}")

    def get_info(self):
        URI = '*****\nDriver: {}\nServer Address: {}\nDatabase Name: {}\nUserID: {}\nPassword: {}\n*****'.format(
            self._driver, self._servername, self._database, self._userid, self._password
        )
        return URI

    @property
    def driver(self):
        return self._driver

    @property
    def server(self):
        return self._servername

    @server.setter
    def server(self, server):
        self._servername = server

    @property
    def database(self):
        return self._database

    @database.setter
    def database(self, database):
        self._database = database

    @property
    def user_id(self):
        return self._userid

    @user_id.setter
    def user_id(self, user_id):
        self._userid = user_id

    @property
    def password(self):
        return self._password

    @password.setter
    def password(self, password):
        self._password = password

    @property
    def engine(self):
        return create_engine(self.dbcon)

    @property
    def connection(self):
        return self.engine.connect()

    @property
    def sql_connection_string(self):
        return f"DRIVER={self._driver};SERVER={self._servername};DATABASE={self._database};UID={self._userid};PWD={self._password}"

    @property
    def dbcon(self):
        return f'mssql+pyodbc:///?odbc_connect={self.sql_connection_string}'

    def connect_engine(self):
        try:
            return self.engine.connect()
        except Exception:
            print("No ODBC Driver detected. Please install Microsoft ODBC Driver 17 for SQL Server from "
                  "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
            webbrowser.open(
                "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
            input("Press Enter to continue...")

    def get_config(self):
        """Method to retrieve the current configuration."""
        return {
            'servername': self.server,
            'database': self.database,
            'userid': self.user_id,
            'password': self.password
        }