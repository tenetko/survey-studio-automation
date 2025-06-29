from datetime import datetime

import pandas as pd
from survey_studio_clients.core.operator_work_time import SurveyStudioOperatorWorkTimeClient

from base_automation import BaseAutomation


class OperatorWorkTimeReportMaker(BaseAutomation):
    PARAMS_NUMBER = 1

    @staticmethod
    def _show_usage_example() -> None:
        print(
            """
        Программу надо запускать одним из двух способов:

        1. Либо с одним параметром - в этом случае программа сформирует отчёт за вчерашний день:

        \t\tpoetry run python get_operator_work_time.py <token>

        2. Либо вообще без параметров - в этом случае программа попросит вас ввести токен и две даты:

        \t\tpoetry run python get_operator_work_time.py
        """
        )

    def _get_raw_data(self) -> pd.DataFrame:
        return self._ss_client.get_dataframe(self._date_from, self._date_to)

    def _get_report_file_name(self) -> str:
        file_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return f"./reports/report_operator_work_time_{self._date_from}_{file_date}.xlsx"

    def run(self) -> None:
        raw_data = self._get_raw_data()
        file_name = self._get_report_file_name()
        raw_data.to_excel(file_name)


if __name__ == "__main__":
    report_maker = OperatorWorkTimeReportMaker(SurveyStudioOperatorWorkTimeClient)
    report_maker.run()
