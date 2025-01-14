import os
import traceback
from pathlib import Path

import DataPuller
import WorkbookHandler
from DataBaseAccessInfo import DataBaseAccessInfo

dpull = DataPuller.DataPuller()
wh = WorkbookHandler.WorkbookHandler()
dbai = DataBaseAccessInfo()

# TODO come up with a way to choose older projects
active_id_df = dpull.active_project_ids()
activeDict = dict.fromkeys(active_id_df['projectid'])
print(active_id_df.to_string())
# quit()

# region manual
# TODO region manual
# date_list = [
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
#     datetime.strptime("2024-06-04", '%Y-%m-%d'),
# ]
#
# # When doing manual, make sure to always put ALL LL projects before their cell equivalents, no matter the date.
# # The dates shown above must be follow the same format as below.
# active_id = {
#     'projectid': [
#         "12842",
#         "12842C",
#         "12847",
#         "12847C",
#         "12849",
#         "12849C",
#     ],
#     'recdate': date_list
# }
# active_id_df = pd.DataFrame(active_id)
# activeDict = dict.fromkeys(active_id_df['projectid'])
# TODO: endregion manual
# endregion manual

# print(active_id_df.to_string())
print(activeDict)
# quit()

def read_excel():
    wh.copy_sheet()
    try:
        if wh.project_code.upper()[-1] == "C":
            wh.active_sheet = "PROJ#C DATE (2)"
        else:
            wh.active_sheet = "PROJ# DATE (2)"
    except Exception as err:
        print(err)
    wh.set_active_sheet_name()

    wh.data = dpull.production_report(wh.project_code, wh.date)
    wh.populate_all()
    wh.populate_expected_loi()

    # need to populate the CPERF sheet, that data must be grabbed from the dpull method. The problem therein lies with
    # changing the DBAI, a process must be put in place to change the DBAI to access the correct database then swap back
    # again

def run_loop():
    try:
        prev = None
        loc = 0
        for project in active_id_df['projectid']:
            projectNumber = project[:5]

            if prev is None or prev != projectNumber or prev != project:
                wh.project_id = projectNumber
                wh.project_code = project

                wh._path = Path(os.environ['src']) / wh.project_id / 'PRODUCTION' / f"{wh.project_id}_Production_Report.xlsm"
            if prev is None or prev != projectNumber:
                wh.check_path()
                while True:
                    try:
                        with open(Path(os.environ['src']) / wh.project_id / 'PRODUCTION' /
                                  f"{wh.project_id}_Production_Report.xlsm", "r+") as file:
                            file.close()
                        break
                    except IOError:
                        input(f"Cannot open file: {wh.project_id}\n\n"
                              f"Please close file then press enter to continue.\n")
                wh.set_workbook()
            prev = projectNumber

            print(f'Completed - {project} - {active_id_df["recdate"][loc]}')
            wh.date = active_id_df['recdate'][loc]
            read_excel()

            if loc >= len(active_id_df['projectid']) - 1 or projectNumber != active_id_df['projectid'][loc + 1][:5]:
                wh.save()
                wh.close()
            loc += 1

        wh.app_quit()
    except:
        print(traceback.format_exc())


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