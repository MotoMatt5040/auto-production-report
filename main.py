import os
import shutil
import DataPuller
import WorkbookHandler
from configparser import ConfigParser
from typing import Union


config_object = ConfigParser()
config_object.read("config.ini")
file_paths = config_object['FILE PATHS']
del config_object


dpull = DataPuller.DataPuller()
wh = WorkbookHandler.WorkbookHandler()


def create_path(projectid):
    if os.path.exists(f"{file_paths['SRC']}{projectid}/PRODUCTION/") == False:
        src = f"{file_paths['SRC']}PRODUCTION/BLANK Production.xlsm"
        os.mkdir(f"{file_paths['SRC']}{projectid}/PRODUCTION/")
        dst = f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_ReportTEST.xlsm"
        shutil.copy(src, dst)
        del src, dst
    elif os.path.exists(f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_ReportTEST.xlsm") == False:
        src = f"{file_paths['SRC']}PRODUCTION/BLANK Production.xlsm"
        dst = f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_ReportTEST.xlsm"
        shutil.copy(src, dst)
        del src, dst
    else:
        print('Directory exists')
        pass

    return f"{file_paths['SRC']}{projectid}/PRODUCTION/{projectid}_Production_ReportTEST.xlsm"


def read_excel(path: str, sheetTitle: str, projectid: Union[int, str]):
    wh.setProjectID(projectid)
    wh.setWorkbook(path)
    wh.setActiveSheet(wh.getWorkbook().sheets[0])
    print(wh.getActiveSheet().name)
    wh.setActiveSheetTitle(sheetTitle)
    print(wh.getActiveSheetTitle())
    wh.copyRows(10)
    wh.populateExpectedLOI()
    wh.copySheet()
    wh.save()
    wh.close()




    # wb = load_workbook(path)
    # sheet1 = wb[wb.sheetnames[0]]
    # print(sheetname)
    # print()
    # print(sheet1.title)
    # print()


# create_path(12523)
active_id_df = dpull.activeProjectIDs()
active_dict = dict.fromkeys(active_id_df['projectid'])

active_dict = {'12523': '12523',
               '12523C': '12523',
               '12517': '12517'
               }

prev = None
for key in active_dict:
    active_dict[key] = key[:5]
    if prev is None or prev != active_dict[key]:
        path = create_path(active_dict[key])

    elif prev == active_dict[key]:
        print(f'RUNNING {active_dict[key]} again')
    prev = active_dict[key]
    read_excel(path, key, active_dict[key])
    # wh.close()
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

