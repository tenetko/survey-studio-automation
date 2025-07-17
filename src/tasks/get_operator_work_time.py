import sys
from datetime import datetime, timedelta
from time import sleep

import pandas as pd
from survey_studio_clients.api_clients.operator_work_time import SurveyStudioOperatorWorkTimeClient
from survey_studio_clients.web_scrapers.daily_counters import DailyCountersPageScraper

from src.tasks.base_automation import BaseAutomation
from src.utils.get_yesterday import get_yesterday_date


class OperatorWorkTimeReportMaker(BaseAutomation):
    MONTH_TO_STRING_FOR_COUNTER = {
        1: "января",
        2: "февраля",
        3: "марта",
        4: "апреля",
        5: "мая",
        6: "июня",
        7: "июля",
        8: "августа",
        9: "сентября",
        10: "октября",
        11: "ноября",
        12: "декабря",
    }

    PARAMS_NUMBER = 2

    def __init__(
        self, client: SurveyStudioOperatorWorkTimeClient, scraper: DailyCountersPageScraper, yesterday: datetime
    ) -> None:
        super().__init__(client)

        _url = sys.argv[2]
        self._scraper_client = scraper(_url)

        self._date_from = self._get_date_as_iso_string(yesterday)
        self._counter = self._get_date_as_survey_studio_counter(yesterday)

    @staticmethod
    def _show_usage_example() -> None:
        print("You have to specify both token and url:\n")
        print('\tpoetry run python src/tasks/get_operator_work_time.py yourtoken123 "https://abc.xyz"')

    def _get_raw_data(self) -> pd.DataFrame:
        return self._ss_client.get_dataframe(self._date_from, self._date_from)

    def _get_date_as_survey_studio_counter(self, dt: datetime) -> str:
        month = dt.month
        day = dt.day

        return f"{day} {self.MONTH_TO_STRING_FOR_COUNTER[month]}"

    def _make_everyday_report(self, df: pd.DataFrame, completes_amount: int) -> pd.DataFrame:
        rows = []
        rows.append(df.iloc[0].iloc[1])

        df = df[2:][:]

        new_header = df.iloc[0]
        df = df[1:]
        df.rename(columns=new_header, inplace=True)
        df.reset_index(inplace=True, drop=True)
        df = df.drop(df.index[-1])

        df = df[df["Наименование"].str.contains("23-071675-18-C", na=False)]
        df = df.reset_index()

        count_oper_name = len(df["Оператор"].unique())
        rows.append(count_oper_name)

        columns_to_change = ["Готов", "Разговор", "Перезвон", "Звонков", "Всего"]
        for c in columns_to_change:
            df[c] = df[c].apply(lambda x: x / 3600)

        worktimes = []
        for _, row in df.iterrows():
            x = row["Готов"] + row["Разговор"] + row["Перезвон"] + row["Звонков"] * 3
            if x > row["Всего"]:
                x = row["Всего"]
            worktimes.append(x)

        df["Рабочее время"] = pd.Series(worktimes)
        work_time = round(df["Рабочее время"].sum(), 2)
        rows.append(work_time)

        rows.append(completes_amount)
        cnt_recruites = df["Успешных"].sum()
        rows.append(cnt_recruites)
        per_questionnaire = round((completes_amount / work_time), 2)
        rows.append(per_questionnaire)
        per_recruits = round((cnt_recruites / work_time), 2)
        rows.append(per_recruits)

        folders = df["Наименование"].unique()
        res_folders = []
        target_value = "2025"

        for el in folders:
            index = el.index(target_value)
            res_folders.append(el[index + 5 :])

        rows.append(", ".join(res_folders))

        rows = [rows]

        return pd.DataFrame(rows)

    def _get_report_file_name(self) -> str:
        file_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return f"./reports/report_operator_work_time_{self._date_from}_{file_date}.xlsx"

    def run(self) -> None:
        raw_data = self._get_raw_data()
        completes_amount = self._scraper_client._get_value_by_counter_name(self._counter)

        report = self._make_everyday_report(raw_data, completes_amount)

        file_name = self._get_report_file_name()
        report.to_excel(file_name)
        print(f"File {file_name} has been successfully saved")


if __name__ == "__main__":
    yesterday = get_yesterday_date()
    if yesterday.weekday() == 6:  # sunday
        for delta in range(2, -1, -1):
            day = yesterday - timedelta(days=delta)
            report_maker = OperatorWorkTimeReportMaker(
                SurveyStudioOperatorWorkTimeClient, DailyCountersPageScraper, day
            )
            report_maker.run()
            if delta != 0:
                print(
                    "Waiting for 62 seconds because Survey Studio API allows only one request to requestoperatorworktime per 60 seconds..."
                )
                sleep(62)
    else:
        report_maker = OperatorWorkTimeReportMaker(
            SurveyStudioOperatorWorkTimeClient, DailyCountersPageScraper, yesterday
        )
        report_maker.run()
