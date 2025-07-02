import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from survey_studio_clients.api_clients.base import SurveyStudioClient


class BaseAutomation:
    DATE_FORMAT = "%Y-%m-%d"
    PARAMS_NUMBER = 0

    def __init__(self, client: SurveyStudioClient) -> None:
        if not self._are_arguments_valid():
            self._show_usage_example()
            sys.exit()

        _token = self._get_token()
        self._date_from = self._get_date_from()
        self._date_to = self._get_date_to()

        self._ss_client = client(_token)

    def _are_arguments_valid(self) -> bool:
        if len(sys.argv) == 1 or self._are_params_provided():
            return True

        return False

    @staticmethod
    def _show_usage_example():
        raise NotImplementedError

    def _are_params_provided(self) -> bool:
        return len(sys.argv) - 1 == self.PARAMS_NUMBER

    def _get_token(self) -> str:
        if self._are_params_provided():
            return sys.argv[1]

        else:
            return input("Введите ваш токен: ")

    def _get_date_from(self) -> str:
        if self._are_params_provided():
            return self._get_yesterday_date()

        else:
            return input("Введите начало периода в формате YYYY-MM-DD hh:mm:ss: ")

    def _get_date_to(self) -> str:
        if self._are_params_provided():
            return self._get_yesterday_date()

        else:
            return input("Введите конец периода в формате YYYY-MM-DD hh:mm:ss: ")

    def _get_yesterday_date(self) -> str:
        date = datetime.now().astimezone(ZoneInfo("Europe/Moscow")) - timedelta(days=1)

        return datetime.strftime(date, self.DATE_FORMAT)

    def _get_report_file_name(self):
        raise NotImplementedError

    def run(self):
        raise NotImplementedError
