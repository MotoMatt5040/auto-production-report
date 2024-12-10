import os
from datetime import date, timedelta


class SQLDictionary:

    def project_id_list(self) -> str:
        """
        Builds active project ID's string for previous day
        :return: String of project ID qry
        """
        if date.today().weekday() == 0:
            time_delta = 3
        else:
            time_delta = 1
        # print(date.today())
        # quit()
        projectids = f"{os.environ['active_project_ids']} '{date.today() - timedelta(time_delta)}'"
        # projectids = "SELECT DISTINCT projectid , recdate FROM tblGPCPHDaily WHERE RecDate >= '2024-12-05'"
        # projectids = f"{os.environ['active project ids']} '2024-05-24'"
        return projectids

    def production_report(self, projectid, date_) -> tuple[str, str, str]:
        """
        Builds production report string for previous day
        :return: String of production report query
        """
        productionReport = f"{os.environ['production_report']}{projectid}{os.environ['production_and']}{date_}{os.environ['production_group_by']}"
        productionReportDispo = f"{os.environ['production_report_dispo']}{projectid}{os.environ['production_report_dispo_and']}{date_}{os.environ['production_report_dispo_single']}"
        productionReportAVGLength = f"{os.environ['production_report_avg_length']}{projectid}{os.environ['production_report_avg_length_and']}{date_}'"

        d = productionReport, productionReportDispo, productionReportAVGLength

        return d
