from datetime import datetime, timedelta
from zoneinfo import ZoneInfo


def get_yesterday_date() -> datetime:
    return datetime.now().astimezone(ZoneInfo("Europe/Moscow")) - timedelta(days=1)
