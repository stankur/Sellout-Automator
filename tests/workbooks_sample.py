from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

import test_rules

def get_sample_distributor_sheet() -> Worksheet:
    source_workbook = load_workbook(
        "/Users/stanleykurniawan/Downloads/BSD MEI 2022.xlsx"
    )
    wanted_sheet_name = "BSD MEI 2022"
    source_sheet = source_workbook[wanted_sheet_name]

    copy_workbook = Workbook()
    copy_sheet = copy_workbook.active

    for column in range(1, test_rules.EXAMPLES_MAX_COLUMNS + 1):
        col = get_column_letter(column)

        for row in range(1, test_rules.EXAMPLES_MAX_ROWS + 1):
            source_cell_content = source_sheet[col + str(row)].value
            copy_sheet[col + str(row)] = source_cell_content

    return copy_sheet
