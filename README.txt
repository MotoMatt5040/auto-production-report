Required setup in order for program to work properly:

    config.ini file with the following fields

        [FILE PATHS]
        SRC = <path>

        [DATABASE ACCESS INFO]
        CORESERVER = <ip>
        CC3SERVER = <ip>
        ARCHIVECC3SERVER = <ip>
        TFSERVER = <ip>
        COREUSER = <username>
        CC3USER = <username>
        ARCHIVECC3USER = <username>
        TFUSER = <username>
        COREPASSWORD = <password>
        CC3PASSWORD = <password>
        ARCHIVECC3PASSWORD = <password>
        TFPASSWORD = <password>
        CALIGULA = <database>

        [SQL CODE]
        ActiveProjectIDs = <qry>
        GPCPH = <qry>
        ProductionReport = <qry>
        ProductionAnd = <qry>
        ProductionGroupBy = <qry>
        ProductionReportDispo = <qry>
        ProductionReportDispoAnd = <qry>
        ProductionReportDispoSingle = <qry>
        ProductionReportAVGLength = <qry>
        ProductioNReportAVGLengthAnd = <qry>

    Adjust variables as needed.

ODBC Driver 17 is required for the application to run properly.
    Download link: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17

if no config file is detected, the application will attempt to create the following directory/file:
    C:\Users\<username>\AppData\Local\AutoProductionReport\config.ini

In order for the application to create the path successfulle, the config.ini folder must be located in your download folder.
    C:\Users\<username>\Downloads\config.ini

If you do not have the config.ini file, manually create one in either directory with the fields above.