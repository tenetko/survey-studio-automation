import sys
from datetime import datetime

from survey_studio_clients.api_clients.base import SurveyStudioClient


class BaseAutomation:
    DATE_FORMAT = "%Y-%m-%d"
    PARAMS_NUMBER = 0

    def __init__(self, client: SurveyStudioClient) -> None:
        if not self._are_arguments_valid():
            self._show_usage_example()
            sys.exit()

        _token = self._get_token()
        self._ss_client = client(_token)

    def _are_arguments_valid(self) -> bool:
        if self._are_params_provided():
            return True

        return False

    def _are_params_provided(self) -> bool:
        return len(sys.argv) - 1 == self.PARAMS_NUMBER

    @staticmethod
    def _show_usage_example():
        raise NotImplementedError

    def _get_token(self) -> str:
        return sys.argv[1]

    def _get_date_from(self) -> str:
        if self._are_params_provided():
            yesterday = self._get_yesterday_date()

            return self._get_date_as_iso_string(yesterday)

        else:
            return input("Введите начало периода в формате YYYY-MM-DD hh:mm:ss: ")

    def _get_date_to(self) -> str:
        if self._are_params_provided():
            yesterday = self._get_yesterday_date()

            return self._get_date_as_iso_string(yesterday)

        else:
            return input("Введите конец периода в формате YYYY-MM-DD hh:mm:ss: ")

    def _get_date_as_iso_string(self, dt: datetime) -> str:
        return datetime.strftime(dt, self.DATE_FORMAT)

    def _get_report_file_name(self):
        raise NotImplementedError

    def run(self):
        raise NotImplementedError
