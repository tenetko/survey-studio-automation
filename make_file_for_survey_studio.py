import json
import sys
from glob import glob

import polars as pl


class SurveyStudioFileMaker:
    TIMEZONES = {
        "UTC +3": "Europe / Moscow",
        "UTC +4": "Europe / Samara",
        "UTC +2": "Europe / Kaliningrad",
        "UTC +5": "Asia / Yekaterinburg",
        "UTC +6": "Asia / Omsk",
        "UTC +7": "Asia / Krasnoyarsk",
        "UTC +8": "Asia / Irkutsk",
        "UTC +9": "Asia / Yakutsk",
        "UTC +10": "Asia / Vladivostok",
        "UTC +11": "Asia / Magadan",
        "UTC +12": "Asia / Kamchatka",
    }

    def __init__(self):
        self.config = self._get_config()

    @staticmethod
    def _get_config() -> dict:
        with open("config.json", "r") as ifile:
            return json.load(ifile)

    def _make_checklist(self) -> dict[str, bool]:
        checklist = {}

        checklist_file_name = ""
        for file_name in glob(f"{self.config['checklist_path']}/*"):
            if "ПРОВЕРКА" not in file_name:
                continue

            checklist_file_name = file_name

        if not checklist_file_name:
            sys.exit("The ПРОВЕРКА file was not found")

        df = pl.read_excel(checklist_file_name, sheet_name="Проверка", read_options={"use_columns": ["number"]})
        for row in df.iter_rows(named=True):
            checklist[str(row["number"])] = True

        return checklist

    def _make_blacklist(self) -> dict[str, bool]:
        blacklist = {}

        blacklist_file_name = ""
        for file_name in glob(f"{self.config['blacklist_path']}/*"):
            if "ЧС" not in file_name:
                continue
            blacklist_file_name = file_name

        if not blacklist_file_name:
            sys.exit("The black list file was not found")

        df = pl.read_excel(blacklist_file_name, read_options={"use_columns": ["Phone"]})
        for row in df.iter_rows(named=True):
            blacklist[str(row["Phone"])] = True

        return blacklist

    def _find_template_file_name(self) -> str:
        template_file_name = ""

        for file_name in glob(f"{self.config['templates_path']}/*"):
            if "template" in file_name or "Temp" in file_name:
                if not self._check_if_file_has_data(file_name):
                    continue

                template_file_name = file_name

        if not template_file_name:
            sys.exit("The template file was not found")

        return template_file_name

    @staticmethod
    def _check_if_file_has_data(file_name: str) -> bool:
        try:
            pl.read_excel(file_name, raise_if_empty=True)
        except pl.exceptions.NoDataError:
            return False

        return True

    @staticmethod
    def _get_source(file_name: str) -> str | None:
        if "_GEN_" in file_name:
            return "GEN_OPER"

        if "_ROBOGEN_RobotCW_" in file_name:
            return "GEN_RobotCW"

        if "_ROBOGEN_TargetAI_" in file_name:
            return "GEN_TARGET"

        sys.Exit("Source was not determined")

    def _make_new_row(self, row: dict, source: str) -> dict:
        return {
            "Number": str(row["tel"]),
            "RegionName": row["obl_name"],
            "OperatorName": row["GrS_name"],
            "TimeDifference": row["UTC_timediff"],
            "Region": self._get_reg_code(row),
            "Operator": str(row["GrS_code"]),
            "CallIntervalBegin": "10:00:00",
            "CallIntervalEnd": "22:00:00",
            "Group": self._get_group(row),
            "CHECK": self._get_check(row),
            "Mark": self._get_mark(row),
            "SOURCE": source,
            "TimeZone": self._get_timezone(row),
        }

    @staticmethod
    def _get_reg_code(row: dict) -> str:
        region_id = int(row["obl_code"])
        if region_id < 10:
            return f"{region_id:0>2}"
        else:
            return str(region_id)

    @staticmethod
    def _get_group(row: dict) -> str:
        return f"{row['obl_name']}_{row['GrS_name']}"

    @staticmethod
    def _get_check(row: dict) -> str:
        return str(row["tel"])[1:]

    @staticmethod
    def _get_mark(row: dict) -> str:
        return f"{row['obl_code']}_{row['GrS_code']}"

    def _get_timezone(self, row: dict) -> str:
        timezone = row["UTC_timediff"]

        return self.TIMEZONES[timezone]

    def _get_output_file_name(self, template_file_name: str) -> str:
        file_name_without_temp = template_file_name.replace("_Temp", "").replace("_template", "")
        clean_file_name = file_name_without_temp.split(".")[1].replace("/", "")

        next_sequence_number = self._get_result_file_name_sequence_number(clean_file_name)

        return f"{clean_file_name}_{next_sequence_number}.xlsx"

    def _get_result_file_name_sequence_number(self, file_name: str) -> str:
        sequence_numbers = []
        existing_result_files = glob(f"{self.config['results_path']}/{file_name}*.xlsx")

        for file_name in existing_result_files:
            sequence_numbers.append(int(file_name.split(".")[-2].split("_")[-1]))

        if len(sequence_numbers) == 0:
            return "001"

        return f"{max(sequence_numbers) + 1:0>3}"

    def _write_dataframe_to_file(self, df: pl.DataFrame, template_file_name: str) -> None:
        output_file_name = self._get_output_file_name(template_file_name)
        output_path = f"{self.config['results_path']}/{output_file_name}"
        df.write_excel(output_path)
        print(f"File {output_path} with {len(df)} records is ready")

    @staticmethod
    def _clean_template_file(template_file_name: str) -> None:
        df = pl.DataFrame([])
        df.write_excel(template_file_name)

    def run(self):
        checklist = self._make_checklist()
        blacklist = self._make_blacklist()

        template_file_name = self._find_template_file_name()
        source = self._get_source(template_file_name)

        df = pl.read_excel(template_file_name)
        print(f"A raw file {template_file_name} with {len(df)} records was opened")

        skipped_counter = 0
        new_rows = []
        for row in df.iter_rows(named=True):
            number = str(row["tel"])

            if number in checklist or number in blacklist:
                print(f"Number {number} was found either in checklist or in blacklist")
                skipped_counter += 1
                continue

            new_row = self._make_new_row(row, source)
            new_rows.append(new_row)

        print(f"{skipped_counter} numbers were skipped because they were found either in checklist or in blacklist")

        df = pl.DataFrame(new_rows)
        print(f"A new dataframe with {len(df)} records was formed")

        split_point = int(len(df) / 2)

        df1 = df.slice(0, split_point)
        self._write_dataframe_to_file(df1, template_file_name)

        df2 = df.slice(split_point)
        self._write_dataframe_to_file(df2, template_file_name)

        self._clean_template_file(template_file_name)
        print(f"A raw file {template_file_name} was cleared")


if __name__ == "__main__":
    maker = SurveyStudioFileMaker()
    maker.run()
