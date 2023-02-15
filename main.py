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

import DataPuller
import WorkbookHandler

dpull = DataPuller.DataPuller()
wh = WorkbookHandler.WorkbookHandler()


def read_excel():
    wh.set_workbook()
    wh.copy_sheet()
    try:
        if wh.get_project_code().upper()[-1] == "C":
            wh.set_active_sheet("PROJ#C DATE (2)")
        else:
            wh.set_active_sheet("PROJ# DATE (2)")
    except Exception as err:
        print(err)
    wh.set_active_sheet_name()
    wh.populate_expected_loi()
    wh.set_data(dpull.production_report(wh.get_project_code()))
    wh.populate_all()
    wh.save()
    wh.close()


active_id_df = dpull.active_project_ids()
activeDict = dict.fromkeys(active_id_df['projectid'])

prev = None
for key in activeDict:
    projectNumber = key[:5]
    wh.set_project_id(projectNumber)
    wh.set_project_code(key)
    wh.set_path(f"{file_paths['SRC']}{wh.get_project_id()}/PRODUCTION/{wh.get_project_id()}_Production_ReportTEST.xlsm")
    if prev is None or prev != activeDict[key]:
        wh.check_path()
    elif prev == activeDict[key]:
        pass
    prev = activeDict[key]
    read_excel()

wh.app_quit()

if __name__ == '__main__':
    print()
