import logging
from datetime import date, timedelta
from logging.handlers import RotatingFileHandler
from configparser import ConfigParser


config_object = ConfigParser()
config_object.read("config.ini")
qry = config_object["SQL CODE"]
del config_object


logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s:%(name)s:%(funcName)s:%(levelname)s:%(message)s')
file_handler = RotatingFileHandler('logs\SQLDictionary.log', mode='a', maxBytes=1024)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.propagate = False


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

        #THIS IS ALL STUFF ON TOP
        #df of 2020 planner, exp l where projectid ==
        #daily inc and avg daily and overall avg, day in review
        #avg mph avg cph, prod2020

        #int data from prod2020





