import sys
from typing import Any

from google.auth.exceptions import RefreshError
from googleapiclient.discovery import Resource
from googleapiclient.errors import HttpError

from src.google_clients.google_auth_client import GoogleAuthClient


class GoogleSheetsClient:
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    SERVICE_NAME = "sheets"
    VALUE_INPUT_OPTION = "RAW"
    VERSION = "v4"

    def __init__(self, spreadsheet_id: str) -> None:
        self.sheet = self.get_service()
        self.values = self.sheet.values()
        self.spreadsheet_id = spreadsheet_id

    def get_service(self) -> Resource:
        params = {
            "scopes": self.SCOPES,
            "service_name": self.SERVICE_NAME,
            "version": self.VERSION,
        }

        try:
            client = GoogleAuthClient(params)
            service = client.build_service()

            return service.spreadsheets()

        except HttpError as e:
            print(f"Error: {e}")
            sys.exit(1)

    def get_sheets(self) -> list:
        data = []

        try:
            result = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
            data = result.get("sheets", "")

        except RefreshError as e:
            print(f"Error: {e}")

        return data

    def get_sheet_id(self, sheet_name: str) -> int | None:
        for sheet in self.get_sheets():
            properties = sheet["properties"]
            if properties["title"] == sheet_name:
                return properties["sheetId"]

        return None

    def read_sheet(self, cells_range: str) -> list[list[Any]]:
        data = []

        try:
            result = self.values.get(spreadsheetId=self.spreadsheet_id, range=cells_range).execute()
            data = result.get("values", [])
        except RefreshError as e:
            print(f"Error: {e}")

        return data

    def write_values(self, values: list[list[Any]], cells_range: str) -> str | None:
        data = [
            {"range": cells_range, "values": values},
        ]

        body = {
            "data": data,
            "valueInputOption": self.VALUE_INPUT_OPTION,
        }

        try:
            result = self.values.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=body,
            ).execute()

            rows = result.get("totalUpdatedRows")
            columns = result.get("totalUpdatedColumns")
            updated_range = result.get("responses")[0]["updatedRange"]

            print(f"{rows} rows and {columns} columns inserted.")

            return updated_range

        except RefreshError as e:
            print(f"Error: {e}")

        except HttpError as e:
            print(f"Error: {e}")

    def append_values(self, values: list[list[Any]], cells_range: str) -> None:
        body = {
            "values": values,
        }

        try:
            result = self.values.append(
                spreadsheetId=self.spreadsheet_id,
                body=body,
                range=cells_range,
                valueInputOption=self.VALUE_INPUT_OPTION,
            ).execute()

        except RefreshError as e:
            print(f"Error: {e}")

        except HttpError as e:
            print(f"Error: {e}")

    def change_sheet(self, requests: list[dict]) -> None:
        body = {"requests": requests}

        try:
            result = self.sheet.batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=body,
            ).execute()

            requests_names = []
            for request in requests:
                requests_names.append(list(request.keys())[0])
            requests_string = ", ".join(requests_names)

            print(f"Requests applied to sheet: {requests_string}.")
            return result

        except RefreshError as e:
            print(f"Error: {e}")

        except HttpError as e:
            print(f"Error: {e}")

    def clear_sheet(self, sheet_id: int) -> None:
        format_requests = [
            {
                "updateCells": {
                    "range": {"sheetId": sheet_id},
                    "fields": "userEnteredValue",
                }
            },
            {
                "updateCells": {
                    "range": {"sheetId": sheet_id},
                    "fields": "userEnteredFormat",
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                    },
                    "properties": {
                        "pixelSize": 100,  # default width for a Google Sheet
                    },
                    "fields": "pixelSize",
                }
            },
        ]

        self.change_sheet(format_requests)

    # [1, 2, 3, 4, 5] --> "A:E"
    # No longer than A:Z
    @staticmethod
    def get_cells_range(row: list[Any]) -> str:
        if len(row) == 0:
            return "A:A"

        if len(row) > 26:
            raise ValueError("Row is too long")

        right_column_index = chr(ord("A") + len(row) - 1)

        return f"A:{right_column_index}"
