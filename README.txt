Required setup in order for program to work properly:

    config.ini file with the following fields

        [FILE PATHS]
        src = <path>
        planner = <path>

        [DATABASE ACCESS INFO]
        coreserver = <ip>
        cc3server = <ip>
        archivecc3server = <ip>
        tfserver = <ip>
        coreuser = <username>
        cc3user = <username>
        archivecc3user = <username>
        tfuser = <username>
        corepassword = <password>
        cc3password = <password>
        archivecc3password = <password>
        tfpassword = <password>
        caligula = <database>

        [SQL CODE]
        active project ids = <qry>
        gpcph = <qry>
        production report = <qry>
        production and = <qry>
        production group by = <qry>
        production report dispo = <qry>
        production report dispo and = <qry>
        production report dispo single = <qry>
        production report avg length = <qry>
        production report avg length and = <qry>

        [ODBC INSTALLED]
        check odbc = 0

    Adjust variables as needed.

ODBC Driver 17 is required for the application to run properly.
    Download link: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17

if no config file is detected, the application will attempt to create the following directory/file:
    C:\Users\<username>\AppData\Local\AutoProductionReport\config.ini

In order for the application to create the path successfulle, the config.ini folder must be located in your download folder.
    C:\Users\<username>\Downloads\config.ini

If you do not have the config.ini file, manually create one in either directory with the fields above.