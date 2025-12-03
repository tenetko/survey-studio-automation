import sys
from datetime import datetime, timedelta
from time import sleep
from zoneinfo import ZoneInfo

import pandas as pd

from survey_studio_clients.api_clients.load_arrow_abbot import SurveyStudioLoadArrowAbbotClient
from base_automation import BaseAutomation

class LoadArrowDailyReportMaker(BaseAutomation):
    PARAMS_NUMBER = 2

    def __init__(self, client: SurveyStudioLoadArrowAbbotClient) -> None:
        super().__init__(client)
        self._project_id = self._get_project_id()
        self._counter_name = "CAWI полные интервью"
    
    def run(self) -> None:
        counter_id = self._ss_client.get_counter_id_by_name(self._project_id, self._counter_name)
        print(counter_id)
    
    def _get_project_id(self) -> str:
        return sys.argv[2]

    def _get_raw_data(self) -> pd.DataFrame:
        raw_data = self._ss_client.get_dataframe(self._project_id)
    
    def _make_everyday_report_abbot(self, df: pd.DataFrame) -> pd.DataFrame:

        if "DB_Аптечная_сетьHidden_Т04" in list(df.columns):
            columnsPharma_temp = ['ID', 'RespExtID', 'UserID', 'UserName', 'UserLgIn', 'IVDate1',
            'IVDate2', 'Phone', 'Result', 'ContactID', 'DB_Город_врачаРегион_врача',
            'DB_CallIntervalEnd', 'DB_CallIntervalBegin', 'DB_UTC',
            'DB_Организация', 'DB_TimeZone', 'DB_Препарат_на_визите',
            'DB_Организация_Род_орг', 'DB_Организация__Улица', 'DB_Внешний_ключ',
            'DB_ID_респондента_T01_hidden', 'DB_Организация__Регион',
            'DB_Связка_Препарат__Специальность', 'DB_Номер_анкетыHidden_Т03',
            'DB_Аптечная_сетьHidden_Т04', 'DB_PHONE_1', 'DB_ОПРОСИТЬ_ДО',
            'DB_Спец_категория_на_визите', 'DB_Организация_Офиц_название',
            'DB_Mark', 'DB_GROUP', 'DB_RESPONDENT_Name', 'DB_Организация__Город',
            'DB_Статус_звонка', 'DB_Выборка_Т02_Hidden',
            'DB_Контакт___Email_Abbott', 'L_qst', 'TYPE', 'Q_100', 'Q_101', 'Q_102',
            'Q_102_7T', 'Q_103', 'Q_104', 'Q_105', 'Q_106', 'Q_107']
            df = df.rename(columns={"DB_Организация__Родительская_организация": "DB_Организация_Род_орг", "DB_Специальностькатегория_на_визите": "DB_Спец_категория_на_визите", "DB_Организация__Официальное_название": "DB_Организация_Офиц_название"})
            df_pharma = df[columnsPharma_temp]
        else:
            columnsDoc_temp = ['ID', 'RespExtID', 'UserID', 'UserName', 'UserLgIn', 'IVDate1',
            'IVDate2', 'Phone', 'Result', 'ContactID', 'DB_Город_врачаРегион_врача',
            'DB_CallIntervalEnd', 'DB_CallIntervalBegin', 'DB_UTC',
            'DB_Специальность_S__0', 'DB_Организация', 'DB_TimeZone',
            'DB_Препарат_на_визите', 'v1', 'DB_Организация__Улица',
            'DB_Внешний_ключ', 'DB_ID_респондента_T01_hidden',
            'DB_Организация__Регион', 'DB_Связка_Препарат__Специальность',
            'DB_Номер_анкетыHidden_Т03', 'DB_Email', 'DB_PHONE_1', 'DB_ОПРОСИТЬ_ДО',
            'v2', 'v3', 'DB_Mark', 'DB_GROUP', 'DB_RESPONDENT_Name',
            'DB_Организация__Город', 'DB_Выборка_Т02_Hidden', 'L_qst', 'TYPE',
            'Q_100', 'Q_101', 'Q_102', 'Q_102_7T', 'Q_103', 'Q_104', 'Q_105',
            'Q_106', 'Q_107']
            df = df.rename(columns={"DB_Организация__Родительская_организация": "v1", "DB_Специальностькатегория_на_визите": "v2", "DB_Организация__Официальное_название": "v3"})
            df_doc = df[columnsDoc_temp]
        
        return df_pharma, df_doc

    def _get_report_file_name(self) -> str:
        file_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name_doc = f"./reports/report_Doc_{file_date}.xlsx"
        file_name_pharma = f"./reports/report_Pharma_{file_date}.xlsx"

        return file_name_doc, file_name_pharma

    def run(self) -> None:
        raw_data = self._get_raw_data()

        file_name = self._get_report_file_name()
        df_pharma.to_excel(file_name_pharma)
        print(f"File {file_name_pharma} has been successfully saved")
        df_doc.to_excel(file_name_doc)
        print(f"File {file_name_doc} has been successfully saved")

        
if __name__ == "__main__":
    report_maker_abbot = LoadArrowDailyReportMaker(SurveyStudioLoadArrowAbbotClient)
    report_maker_abbot.run()