import os
from configparser import ConfigParser
from datetime import date, timedelta
from pathlib import Path

config_object = ConfigParser()
config_path = Path(f"C:/Users/{os.getlogin()}/AppData/Local/AutoProductionReport/config.ini")
config_object.read(config_path)
qry = config_object["SQL CODE"]
del config_object


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
        projectids = f"{qry['active project ids']} '{date.today() - timedelta(1)}'"
        # projectids = "SELECT DISTINCT projectid , recdate FROM tblGPCPHDaily WHERE RecDate = '2024-05-28'"
        # projectids = f"{qry['active project ids']} '2024-05-24'"
        # print(projectids)
        # quit()
        return projectids

    def production_report(self, projectid, date_) -> tuple[str, str, str]:
        """
        Builds production report string for previous day
        :return: String of production report qry
        """
        productionReport = f"{qry['production report']}{projectid}{qry['production and']}{date_}{qry['production group by']}"
        productionReportDispo = f"{qry['production report dispo']}{projectid}{qry['production report dispo and']}{date_}{qry['production report dispo single']}"
        productionReportAVGLength = f"{qry['production report avg length']}{projectid}{qry['production report avg length and']}{date_}'"

        d = productionReport, productionReportDispo, productionReportAVGLength

        # for item in d:
        #     print(item)
        # quit()

        return d
