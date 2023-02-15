from datetime import date, timedelta
from configparser import ConfigParser
import os


config_object = ConfigParser()
config_path = f"C:/Users/{os.getlogin()}/AppData/Local/AutoProductionReport/config.ini"
config_object.read(config_path)
qry = config_object["SQL CODE"]
del config_object


class SQLDictionary:

    def project_id_list(self) -> str:
        """
        Builds active project ID's string for previous day
        :return: String of project ID qry
        """
        projectids = f"{qry['active project ids']} '{str(date.today() - timedelta(1))}'"
        return projectids

    def gpcph(self) -> str:
        """
        Builds GPCPH string for previous day
        :return: String of GPCPH qry
        """
        gpcph = f"{qry['gpcph']} '{str(date.today() - timedelta(1))}'"
        return gpcph

    def production_report(self, projectid) -> tuple[str, str, str]:
        """
        Builds production report string for previous day
        :return: String of production report qry
        """
        productionReport = f"{qry['production report']}{projectid}{qry['production and']}{date.today() - timedelta(1)}{qry['production group by']}"
        productionReportDispo = f"{qry['production report dispo']}{projectid}{qry['production report dispo and']}{date.today() - timedelta(1)}{qry['production report dispo single']}"
        productionReportAVGLength = f"{qry['production report avg length']}{projectid}{qry['production report avg length and']}{date.today() - timedelta(1)}'"

        return productionReport, productionReportDispo, productionReportAVGLength
