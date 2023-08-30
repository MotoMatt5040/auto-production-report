import os
import traceback
import webbrowser
import pandas as pd
from datetime import datetime
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
    wh.set_data(dpull.production_report(wh.get_project_code(), wh.get_date()))
    wh.populate_all()


# TODO come up with a way to choose older projects
active_id_df = dpull.active_project_ids()
activeDict = dict.fromkeys(active_id_df['projectid'])
# print(active_id_df)
# quit()
#
# print(active_id_df.to_string())
# print(activeDict)
# quit()

# print(active_id_df.to_string())
# print(activeDict)
# TODO REMOVE BELOW, this was a 1 off for a specific project
# date_list = [
#     datetime.strptime('2023-07-21', '%Y-%m-%d'),
#     datetime.strptime('2023-07-22', '%Y-%m-%d'),
#     datetime.strptime('2023-07-21', '%Y-%m-%d'),
#     datetime.strptime('2023-07-22', '%Y-%m-%d'),
# ]
#
# active_id = {
#     'projectid': ['12644', '12644', '12644C', '12644C'],
#     'recdate': date_list
# }
# active_id_df = pd.DataFrame(active_id)
# activeDict = dict.fromkeys(active_id_df['projectid'])
#
# print(active_id_df.to_string())
# print(activeDict)
# quit()
#
# # 20210809	20210815

try:

    prev = None
    loc = 0
    for project in active_id_df['projectid']:
        projectNumber = project[:5]

        if prev is None or prev != projectNumber or prev != project:
            wh.set_project_id(projectNumber)
            wh.set_project_code(project)
            wh.set_path(f"{file_paths['SRC']}{wh.get_project_id()}/PRODUCTION/{wh.get_project_id()}_Production_Report.xlsm")

        if prev is None or prev != projectNumber:
            wh.check_path()
            wh.set_workbook()
        prev = projectNumber

        wh.set_date(active_id_df['recdate'][loc])
        read_excel()

        if loc >= len(active_id_df['projectid']) - 1 or projectNumber != active_id_df['projectid'][loc+1][:5]:
            wh.save()
            wh.close()
        loc += 1


    wh.app_quit()
except:
    print(traceback.format_exc())
    input("Press enter to continue")

if __name__ == '__main__':
    print()
