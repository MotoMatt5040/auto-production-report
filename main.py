import DataPuller
import WorkbookHandler
from configparser import ConfigParser


config_object = ConfigParser()
config_object.read("config.ini")
file_paths = config_object['FILE PATHS']
del config_object


dpull = DataPuller.DataPuller()
wh = WorkbookHandler.WorkbookHandler()


def read_excel():
    wh.set_workbook()
    try:
        if wh.get_project_code().upper()[-1] == "C":
            wh.set_active_sheet("PROJ#C DATE")
            print('cell')
        else:
            wh.set_active_sheet("PROJ# DATE")
            print('ll')
    except Exception as err:
        print(err)
        wh.copy_sheet()
        if wh.get_project_code().upper()[-1] == "C":
            wh.set_active_sheet("PROJ#C DATE")
            print('cell')
        else:
            wh.set_active_sheet("PROJ# DATE")
            print('ll')
    print(wh.get_active_sheet_name())
    wh.set_active_sheet_name(f"{wh.get_project_code()}")
    wh.populate_expected_loi()
    print(wh.get_active_sheet())
    wh.set_data(dpull.production_report(wh.get_project_code()))
    wh.populate_all()
    # wh.populate_production_data()
    # print(dpull.production_report(wh.get_project_code())[0].to_string())
    # print(dpull.production_report(wh.get_project_code())[1].to_string())
    # print(dpull.production_report(wh.get_project_code())[2].to_string())
    wh.save()
    wh.close()


active_id_df = dpull.active_project_ids()
activeDict = dict.fromkeys(active_id_df['projectid'])


# activeDict = {'12523': '12523',
#                '12523C': '12523'
#                }

# TODO Make sure wh.expectedloi works properly


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
    # print(dpull.productionReport(wh.getProjectCode()).to_string())
    read_excel()



print(activeDict)
active_id_list_details = list(active_id_df['projectid'])
active_id_list = [projectid[:5] for projectid in active_id_list_details]


if __name__ == '__main__':
    print()

