import traceback
import DataPuller
import WorkbookHandler
import checkodbc

file_paths = checkodbc.file_paths
dpull = DataPuller.DataPuller()
wh = WorkbookHandler.WorkbookHandler()

# TODO come up with a way to choose older projects
active_id_df = dpull.active_project_ids()
activeDict = dict.fromkeys(active_id_df['projectid'])

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

def run_loop():
    try:
        prev = None
        loc = 0
        for project in active_id_df['projectid']:
            projectNumber = project[:5]

            if prev is None or prev != projectNumber or prev != project:
                wh.set_project_id(projectNumber)
                wh.set_project_code(project)

                wh.set_path(
                    f"{file_paths['SRC']}{wh.get_project_id()}/PRODUCTION/{wh.get_project_id()}_Production_Report.xlsm")
            if prev is None or prev != projectNumber:
                wh.check_path()
                wh.set_workbook()
            prev = projectNumber
            while True:
                try:
                    with open(f"{file_paths['SRC']}{wh.get_project_id()}/PRODUCTION/"
                              f"{wh.get_project_id()}_Production_Report.xlsm", "r+") as file:
                        break
                except IOError:
                    input(f"Cannot open file: {wh.get_project_id()}\n\n"
                          f"Please close file then press enter to continue.\n")

            wh.set_date(active_id_df['recdate'][loc])
            read_excel()

            if loc >= len(active_id_df['projectid']) - 1 or projectNumber != active_id_df['projectid'][loc + 1][:5]:
                wh.save()
                wh.close()
            loc += 1

        wh.app_quit()
    except:
        print(traceback.format_exc())
        input("Press enter to continue")


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