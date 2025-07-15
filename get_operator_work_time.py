from datetime import datetime

import pandas as pd
from survey_studio_clients.api_clients.operator_work_time import SurveyStudioOperatorWorkTimeClient
from survey_studio_clients.web_scrapers.daily_counters import DailyCountersPageScraper

from base_automation import BaseAutomation
import json



class OperatorWorkTimeReportMaker(BaseAutomation):
    PARAMS_NUMBER = 1

    def __init__(self, client: SurveyStudioOperatorWorkTimeClient, scraper_client: DailyCountersPageScraper) -> None:
        self._config = self._get_config()
        self._ss_client = client(self._config["token"])
        self._scraper_client = DailyCountersPageScraper(self._config["quota_url"])
        self._date_from = self._config["date_from"]
        self._date_to = self._config["date_to"]

    def _get_config(self):
        with open("config.json") as ifile:
            return json.load(ifile)

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

    def _make_everyday_report(self, df: pd.DataFrame) -> None:
        counter_name = self._scraper_client.get_daily_counter_name(datetime(2025, 7, 9))
        value_comp = self._scraper_client.get_value_by_counter_name(counter_name)

        rows = []
        rows.append(df.iloc[0].iloc[1])
        
        df = df[2:][:]

        new_header = df.iloc[0]
        df = df[1:]
        df.rename(columns=new_header, inplace=True)
        df.reset_index(inplace=True, drop=True)
        df = df.drop(df.index[-1])

        df = df[df['Наименование'].str.contains('23-071675-18-C', na=False)]
        df = df.reset_index()

        count_oper_name = len(df['Оператор'].unique())
        rows.append(count_oper_name)

        columns_to_change = ['Готов','Разговор','Перезвон','Звонков', 'Всего']
        for c in columns_to_change:
            df[c] = df[c].apply(lambda x: x / 3600)

        worktimes = []
        for _, row in df.iterrows():
            x = row['Готов'] + row['Разговор'] + row['Перезвон'] + row['Звонков']*3
            if x > row['Всего']:
                x = row['Всего']
            worktimes.append(x)

        df['Рабочее время'] = pd.Series(worktimes)
        work_time = round(df['Рабочее время'].sum(),2 )
        rows.append(work_time)
        
        rows.append(value_comp)
        cnt_recruites = df['Успешных'].sum()        
        rows.append(cnt_recruites)
        per_questionnaire = round((value_comp/work_time), 2)
        rows.append(per_questionnaire)
        per_recruits = round((cnt_recruites/work_time), 2)
        rows.append(per_recruits)

        folders = df_worktime['Наименование'].unique()
        res_folders = []
        target_value = '2025'

        for el in folders:
            index = el.index(target_value)
            res_folders.append(el[index + 5:])

        columns.append(", ".join(res_folders))  

        rows = [rows]

        return pd.DataFrame(rows)
    
    def run(self) -> None:
        
       
      

        raw_data = self._get_raw_data()
        report = self._make_everyday_report(raw_data)
        print(report)

        file_name = self._get_report_file_name()
        raw_data.to_excel(file_name)
        

    


if __name__ == "__main__":
    report_maker = OperatorWorkTimeReportMaker(SurveyStudioOperatorWorkTimeClient, DailyCountersPageScraper)
    report_maker.run()
