from turtle import st
from openpyxl import load_workbook
import helpers


class ExcelFile:
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        self.workbook = load_workbook(excel_file_path)
        self.helper = helpers.Helper(self.workbook[sheet_label])
        self.id_value_column = self.helper.create_value_column(id_label)
        self.data_column = self.helper.get_column(data_label)


class SourceFile(ExcelFile):
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        super().__init__(excel_file_path, sheet_label, id_label, data_label)

    def get_all_ids_and_data(self):
        pass


class DestinationFile(ExcelFile):
    def __init__(
        self, excel_file_path: str, sheet_label: str, id_label: str, data_label: str
    ):
        super().__init__(excel_file_path, sheet_label, id_label, data_label)

    def update_values(self):
        pass
