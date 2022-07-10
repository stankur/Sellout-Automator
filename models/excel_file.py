from typing import Any, List, Tuple
from openpyxl import load_workbook
import models.helpers as helpers


class ExcelFile:
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        self.excel_file_path = excel_file_path
        self.workbook = load_workbook(excel_file_path)
        self.sheet = self.workbook[sheet_label]
        self.helper = helpers.Helper(self.workbook[sheet_label])
        self.id_value_column = self.helper.create_value_column(id_label)
        self.data_column = self.helper.get_column(data_label)

    def save(self) -> None:
        self.workbook.save(self.excel_file_path)


class SourceFile(ExcelFile):
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        super().__init__(excel_file_path, sheet_label, id_label, data_label)

    def get_all_ids_and_data(self) -> List[Tuple[Any, Any]]:
        all_ids_and_data = []

        for row in range(
            self.id_value_column.get_start_cell_row(),
            self.id_value_column.get_end_cell_row() + 1,
        ):
            all_ids_and_data.append(
                (
                    self.id_value_column.get_value_at(row),
                    self.sheet[self.data_column + str(row)].value,
                )
            )

        return all_ids_and_data


class DestinationFile(ExcelFile):
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        super().__init__(excel_file_path, sheet_label, id_label, data_label)

    def update_values(self, ids_and_data: List[Tuple[Any, Any]]) -> None:
        for id_and_datum in ids_and_data:
            id = id_and_datum[0]
            datum = id_and_datum[1]

            to_be_replace_row = self.id_value_column.get_row_of(id)
            self.sheet[self.data_column + str(to_be_replace_row)].value = datum


def update(source_file: SourceFile, destination_file: DestinationFile) -> None:
    destination_file.update_values(source_file.get_all_ids_and_data())
    destination_file.save()
