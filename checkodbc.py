import os
import webbrowser
from configparser import ConfigParser
import shutil

config_path = f"C:/Users/{os.getlogin()}/AppData/Local/AutoProductionReport/"
if not os.path.exists(config_path):
    os.mkdir(config_path)
if not os.path.exists(f"{config_path}config.ini"):
    shutil.copyfile(f'C:/Users/{os.getlogin()}/Downloads/config.ini', f"{config_path}config.ini")

config_object = ConfigParser()
config_object.read(f"{config_path}config.ini")
file_paths = config_object['FILE PATHS']
check_odbc = config_object['ODBC INSTALLED']

if check_odbc['check odbc'] == '0':
    print("*" * 20)
    print('ODBC Driver 17 must be installed.')
    verify = False
    while not verify:
        option = input("ODBC Driver 17 already installed? (Y/N)")
        if type(option) != str:
            print('Invalid input.')
        elif option.lower() != 'y' and option.lower() != 'n':
            print(option.lower() == 'y')
            print('Invalid input.')
        elif option.lower() == 'y':
            config_object.set('ODBC INSTALLED', 'check odbc', '1')
            with open(f"{config_path}config.ini", 'w') as configfile:
                config_object.write(configfile)
            verify = True
        elif option.lower() == 'n':
            input("Please install Microsoft ODBC Driver 17 for SQL Server from:\n "
                  "    https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17\n\n"
                  "Press enter to open webpage...")
            webbrowser.open(
                "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#version-17")
            input('Press enter to continue once ODBC Driver 17 is finished installing...')
            config_object.set('ODBC INSTALLED', 'check odbc', '1')
            with open(f"{config_path}config.ini", 'w') as configfile:
                config_object.write(configfile)
            verify = True
del config_object