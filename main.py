import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.messagebox import showerror
import os
import shutil
from openpyxl import load_workbook
import SQLDictionary
import DataPuller
from configparser import ConfigParser


config_object = ConfigParser()
config_object.read("config.ini")
qry_info = config_object["SQL CODE"]
file_paths = config_object['FILE PATHS']


sqld = SQLDictionary.SQLDictionary()
dpull = DataPuller.DataPuller()


def create_path(projectid):
    if os.path.exists(f"{file_paths['SRC']}{projectid}/PRODUCTION/") == False:
        src = f"{file_paths['SRC']}PRODUCTION/BLANK Production.xlsx"
        os.mkdir(f"{file_paths['SRC']}{projectid}/PRODUCTION/")
        dst = f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_Report.xlsx"
        shutil.copy(src, dst)
        del src, dst
        # print(f"{projectid} created - Production")
    elif os.path.exists(f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_Report.xlsx") == False:
        src = f"{file_paths['SRC']}PRODUCTION/BLANK Production.xlsx"
        dst = f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_Report.xlsx"
        shutil.copy(src, dst)
        del src, dst
        # print(f"{projectid} created")
    else:
        # print(f"{projectid} exists")
        pass

    return f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_Report.xlsx"


# def read_excel(path: str):
#     workbook = load_workbook(path)
#     return workbook



create_path(12522)
active_id_df = dpull.pull_data(sqld.project_id_list(qry=qry_info["ActiveProjectIDs"]))
active_dict = dict.fromkeys(active_id_df['projectid'])

prev = None
for key in active_dict:
    active_dict[key] = key[:5]
    if prev is None or prev != active_dict[key]:
        path = create_path(active_dict[key])
    elif prev == active_dict[key]:
        print(f'RUNNING {active_dict[key]} again')
    prev = active_dict[key]
    # print(path)


# keys = list(active_dict.keys())
# prev = None
# for key in active_dict:
#     if prev is not None or prev != key:
#         create_path(active_dict[key])
#     prev = key

print(active_dict)
active_id_list_details = list(active_id_df['projectid'])
active_id_list = [projectid[:5] for projectid in active_id_list_details]


if __name__ == '__main__':
    print()

