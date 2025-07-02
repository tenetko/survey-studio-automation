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

    def _make_everyday_report(self, df: pd.DataFrame) -> None:
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
        print(worktimes)
        print(df['Рабочее время'])
        work_time = df['Рабочее время'].sum()
        rows.append(work_time)

        cnt_recruites = df['Успешных'].sum()        
        rows.append(cnt_recruites)
        rows = [rows]

        return pd.DataFrame(rows)
    
    def run(self) -> None:
        raw_data = self._get_raw_data()
        report = self._make_everyday_report(raw_data)
        print(report)

        # file_name = self._get_report_file_name()
        # raw_data.to_excel(file_name)
        

    


if __name__ == "__main__":
    report_maker = OperatorWorkTimeReportMaker(SurveyStudioOperatorWorkTimeClient)
    report_maker.run()
