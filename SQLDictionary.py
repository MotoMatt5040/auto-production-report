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
        projectids = f"{qry['ActiveProjectIDs']} '{str(date.today() - timedelta(1))}'"
        return projectids

    def gpcph(self) -> str:
        """
        Builds GPCPH string for previous day
        :return: String of GPCPH qry
        """
        gpcph = f"{qry['GPCPH']} '{str(date.today() - timedelta(1))}'"
        return gpcph

    def production_report(self, projectid) -> tuple[str, str, str]:
        """
        Builds production report string for previous day
        :return: String of production report qry
        """
        productionReport = f"{qry['ProductionReport']}{projectid}{qry['ProductionAnd']}{date.today() - timedelta(1)}{qry['ProductionGroupBy']}"
        productionReportDispo = f"{qry['ProductionReportDispo']}{projectid}{qry['ProductioNReportDispoAnd']}{date.today() - timedelta(1)}{qry['ProductionReportDispoSingle']}"
        productionReportAVGLength = f"{qry['ProductionReportAVGLength']}{projectid}{qry['ProductionReportAVGLengthAnd']}{date.today() - timedelta(1)}'"

        return productionReport, productionReportDispo, productionReportAVGLength
