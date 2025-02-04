import os
from datetime import date, timedelta, datetime
from utils.logger_config import logger
date_format = "%Y-%m-%d"


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
        projectids = f"{os.environ['active_project_ids']} '{date.today() - timedelta(time_delta)}'"
        # projectids = "SELECT DISTINCT projectid , recdate FROM tblGPCPHDaily WHERE RecDate >= '2024-07-29' and RecDate <= '2024-07-31' and projectid = '12886C'"#" or projectid = '12886'"
        # projectids = "SELECT DISTINCT projectid , recdate FROM tblGPCPHDaily WHERE RecDate >= '2025-01-28' and RecDate <= '2025-01-31' and projectid = '13042C' or projectid = '13042'"
        # projectids = "SELECT DISTINCT projectid , recdate FROM tblGPCPHDaily WHERE RecDate = '2025-01-10'"
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

    def voxco_sample_data(self, database: str, start_date: datetime) -> str:
        """"""
        end_date = start_date + timedelta(hours=24)

        qry = f"""
            SELECT {os.environ['voxco_columns']}
            FROM [{database}]{os.environ['his_table']}
            WHERE {os.environ['call_date']} >= '{start_date} 10:00' 
            AND {os.environ['call_date']} < '{end_date} 10:00'
            AND {os.environ['voxco_where']}
            """

        return qry

    def voxco_prel_sample_data(self, database: str, start_date: datetime) -> str:
        """"""
        end_date = start_date + timedelta(hours=24)

        qry = f"""
            SELECT {os.environ['prel_columns']}
            FROM [{database}]{os.environ['his_table']}
            INNER JOIN [{database}]{os.environ['res_table']}
            ON {os.environ['prel_on']}
            WHERE {os.environ['call_date']} >= '{start_date} 10:00' 
            AND {os.environ['call_date']} < '{end_date} 10:00'
            AND {os.environ['rpsq']}
            """

        return qry

    def voxco_unscanned_sample_data(self, database: str, start_date: datetime) -> str:

        end_date = start_date + timedelta(hours=24)

        qry = f"""SELECT {os.environ['unscanned_columns']}
            FROM [{database}]{os.environ['his_table']} AS Historic
            LEFT JOIN [{database}]{os.environ['res_table']} AS Response
            ON {os.environ['prel_on']}
            AND Response.{os.environ['rpsq']}
            WHERE ({os.environ['unscanned_where']})
            AND {os.environ['a4s']}
            AND {os.environ['case']}
            AND {os.environ['call_date']} >= '{start_date} 10:00' 
            AND {os.environ['call_date']} < '{end_date} 10:00'
        """

        return qry
