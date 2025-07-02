import sys
from datetime import datetime, timedelta
from http import HTTPStatus
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup


class DailyCountersParser:
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

    def __init__(self, url: str) -> None:
        self.url = url

    def _parse(self, counter_name_to_find: str) -> str | None:
        res = requests.get(self.url)
        if res.status_code != HTTPStatus.OK:
            raise requests.HTTPError

        rows = BeautifulSoup(res.text, "lxml").find_all("tr")
        for row in rows:
            cells = row.find_all("td")
            if len(cells) != 4:
                continue

            counter_name = cells[0].get_text(strip=True)
            if counter_name != counter_name_to_find:
                continue

            return cells[2].get_text(strip=True)

        return None

    def _get_counter_name(self, date: datetime) -> str:
        month = date.month
        day = date.day

        return f"{day} {self.MONTH_TO_STRING[month]}"

    def run(self):
        yesterday = datetime.now().astimezone(ZoneInfo("Europe/Moscow")) - timedelta(days=1)
        counter_name = self._get_counter_name(yesterday)
        value = self._parse(counter_name)


if __name__ == "__main__":
    parser = DailyCountersParser(sys.argv[1])
    parser.run()
