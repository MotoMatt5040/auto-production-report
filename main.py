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
    wh.setWorkbook()
    try:
        if wh.getProjectCode().upper()[-1] == "C":
            wh.setActiveSheet("PROJ#C DATE")
            print('cell')
        else:
            wh.setActiveSheet("PROJ# DATE")
            print('ll')
    except Exception as err:
        print(err)
        wh.copySheet()
        if wh.getProjectCode().upper()[-1] == "C":
            wh.setActiveSheet("PROJ#C DATE")
            print('cell')
        else:
            wh.setActiveSheet("PROJ# DATE")
            print('ll')
    print(wh.getActiveSheetName())
    wh.setActiveSheetName(f"{wh.getProjectCode()}")
    print(wh.getActiveSheet())
    wh.copyRows(5)
    wh.save()
    wh.close()


active_id_df = dpull.activeProjectIDs()
# active_dict = dict.fromkeys(active_id_df['projectid'])


activeDict = {'12523': '12523',
               '12523C': '12523'
               }


prev = None
for key in activeDict:
    projectNumber = key[:5]
    wh.setProjectID(projectNumber)
    wh.setProjectCode(key)
    wh.setPath(f"{file_paths['SRC']}{wh.getProjectID()}/PRODUCTION/{wh.getProjectID()}_Production_ReportTEST.xlsm")
    if prev is None or prev != activeDict[key]:
        wh.checkPath()
    elif prev == activeDict[key]:
        pass
    prev = activeDict[key]
    read_excel()


print(activeDict)
active_id_list_details = list(active_id_df['projectid'])
active_id_list = [projectid[:5] for projectid in active_id_list_details]


if __name__ == '__main__':
    print()

