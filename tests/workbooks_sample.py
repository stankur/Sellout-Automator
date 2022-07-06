from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from tests import test_rules


def _get_sheet_copy(path: str, sheet_name: str) -> Worksheet:
    source_workbook = load_workbook(path, data_only=True)
    wanted_sheet_name = sheet_name
    source_sheet = source_workbook[wanted_sheet_name]

    copy_workbook = Workbook()
    copy_sheet = copy_workbook.active

    for column in range(1, test_rules.EXAMPLES_MAX_COLUMNS + 1):
        col = get_column_letter(column)

        for row in range(1, test_rules.EXAMPLES_MAX_ROWS + 1):
            source_cell_content = source_sheet[col + str(row)].value
            copy_sheet[col + str(row)] = source_cell_content

    return copy_sheet


def get_sample_distributor_description_as_id_sheet() -> Worksheet:
    return _get_sheet_copy(
        "/Users/stanleykurniawan/Downloads/BSD MEI 2022.xlsx", "BSD MEI 2022"
    )


def get_sample_distributor_code_as_id_sheet() -> Worksheet:
    return _get_sheet_copy(
        "/Users/stanleykurniawan/Downloads/saputra alsut mei 2022.xlsx", "Sheet1"
    )


def get_sample_master_sheet() -> Worksheet:
    return _get_sheet_copy(
        "/Users/stanleykurniawan/Downloads/MASTER NY JANUARI 2022.xlsx", "Upload"
    )
