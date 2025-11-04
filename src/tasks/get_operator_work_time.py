import sys
from datetime import datetime, timedelta
from time import sleep
from zoneinfo import ZoneInfo

import pandas as pd
from survey_studio_clients.api_clients.operator_work_time import SurveyStudioOperatorWorkTimeClient
from survey_studio_clients.web_scrapers.daily_counters import DailyCountersPageScraper

from src.google_clients.google_sheets_client import GoogleSheetsClient
from src.settings import OPERATOR_WORK_TIME_SPREADSHEET_ID, QUOTA_LINK
from src.tasks.base_automation import BaseAutomation
from src.types.google_sheets import CellsRange, RepeatCellRequest
from src.utils.get_yesterday import get_yesterday_date


class OperatorWorkTimeReportMaker(BaseAutomation):
    BORDER = {"style": "SOLID"}

    CELL_FORMAT = {
        "borders": {
            "top": BORDER,
            "bottom": BORDER,
            "right": BORDER,
        },
    }

    MONTH_TO_STRING = {
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

    def __init__(self, client: SurveyStudioOperatorWorkTimeClient, yesterday: datetime) -> None:
        super().__init__(client)

        self._project_name = sys.argv[2]
        self._scraper = DailyCountersPageScraper(QUOTA_LINK)

        self._date_from = self._get_date_as_iso_string(yesterday)
        self._counter = self._get_date_as_survey_studio_counter(yesterday)

        self._sheets = GoogleSheetsClient(OPERATOR_WORK_TIME_SPREADSHEET_ID)
        self._sheet_name = "Отчёт для КМ"

    @staticmethod
    def _show_usage_example() -> None:
        print("You have to specify your token and project name:\n")
        print("\tpoetry run python src/tasks/get_operator_work_time.py yourtoken123 23-012345-67-C")

    def _get_raw_data(self) -> pd.DataFrame:
        return self._ss_client.get_dataframe(self._date_from, self._date_from)

    def _does_report_already_exist(self, raw_data: pd.DataFrame) -> bool:
        date = self._get_date_for_google_sheets(raw_data.iloc[0].iloc[1])
        existing_rows = self._sheets.read_sheet(self._sheet_name)

        for row in existing_rows:
            if row[0] == date:
                return True

        return False

    def _get_date_as_survey_studio_counter(self, dt: datetime) -> str:
        month = dt.month
        day = dt.day

        return f"{day} {self.MONTH_TO_STRING[month]}"

    def _make_everyday_report(self, df: pd.DataFrame, completes_amount: int) -> tuple[list, pd.DataFrame]:
        rows = []
        date = self._get_date_for_google_sheets(df.iloc[0].iloc[1])
        rows.append(date)

        df = df[2:][:]

        new_header = df.iloc[0]
        df = df[1:]
        df.rename(columns=new_header, inplace=True)
        df.reset_index(inplace=True, drop=True)
        df = df.drop(df.index[-1])

        df = df[df["Наименование"].str.contains(self._project_name, na=False)]
        df = df.reset_index()

        columns_to_change = ["Готов", "Разговор", "Перезвон", "Звонков", "Всего"]
        for c in columns_to_change:
            df[c] = df[c].apply(lambda x: x / 3600)

        count_oper_name = len(df["Оператор"].unique())
        rows.append(count_oper_name)

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

        return rows, df

    def _get_date_for_google_sheets(self, dt: str) -> str:
        dt = datetime.strptime(dt, "%d.%m.%Y %H:%M")
        month = self.MONTH_TO_STRING[dt.month]

        return f"{dt.day} {month} {dt.year} г."

    def _get_report_file_name(self) -> str:
        file_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        return f"./reports/report_operator_work_time_{self._date_from}_{file_date}.xlsx"

    def _get_last_row_index(self) -> int:
        values = self._sheets.read_sheet(self._sheet_name)

        return len(values) - 1

    def run(self) -> None:
        raw_data = self._get_raw_data()

        if self._does_report_already_exist(raw_data):
            print(f"The report for {self._date_from} already exists")
            exit()

        completes_amount = self._scraper.get_value_by_counter_name(self._counter)
        rows, dataframe = self._make_everyday_report(raw_data, completes_amount)

        file_name = self._get_report_file_name()
        dataframe.to_excel(file_name)

        print(f"File {file_name} has been successfully saved")

        self._sheets.append_values([rows], self._sheet_name)

        sheet_id = self._sheets.get_sheet_id(self._sheet_name)
        last_row_index = self._get_last_row_index()

        cells_range = CellsRange(sheet_id, last_row_index, last_row_index + 1, 0, 8)
        request = RepeatCellRequest(cells_range, self.CELL_FORMAT)
        format_requests = [request.to_dict()]

        self._sheets.change_sheet(format_requests)

        print(f"A new row for {self._date_from} has been successfully added to the Google Sheet")


if __name__ == "__main__":
    yesterday = get_yesterday_date()
    # yesterday = datetime(2025, 9, 7).astimezone(ZoneInfo("Europe/Moscow"))

    if yesterday.weekday() == 6:  # sunday
        for delta in range(2, -1, -1):
            day = yesterday - timedelta(days=delta)
            report_maker = OperatorWorkTimeReportMaker(SurveyStudioOperatorWorkTimeClient, day)
            report_maker.run()
            if delta != 0:
                print(
                    "Waiting for 62 seconds because Survey Studio API allows only one request to requestoperatorworktime per 60 seconds..."
                )
                sleep(62)
    else:
        report_maker = OperatorWorkTimeReportMaker(SurveyStudioOperatorWorkTimeClient, yesterday)
        report_maker.run()
