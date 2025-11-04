from dataclasses import dataclass


@dataclass
class CellsRange:
    sheet_id: int
    start_row_index: int
    end_row_index: int
    start_column_index: int
    end_column_index: int

    def to_dict(self) -> dict:
        return {
            "sheetId": self.sheet_id,
            "startRowIndex": self.start_row_index,
            "endRowIndex": self.end_row_index,
            "startColumnIndex": self.start_column_index,
            "endColumnIndex": self.end_column_index,
        }


@dataclass
class RepeatCellRequest:
    range: CellsRange
    format: dict

    def to_dict(self) -> dict:
        return {
            "repeatCell": {
                "range": self.range.to_dict(),
                "fields": self.get_fields_string(),
                "cell": {
                    "userEnteredFormat": self.format,
                },
            }
        }

    def get_fields_string(self) -> str:
        fields = ", ".join(self.format.keys())
        fields_string = f"userEnteredFormat({fields})"

        return fields_string
