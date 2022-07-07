import pytest
from openpyxl import load_workbook
import models.excel_file as excel_file
import workbooks_sample

def test_update():
    # saved at workbook_copy.xlsx
    workbooks_sample.get_sample_distributor_code_as_id_sheet()

    # saved at workbook_copy_1.xlsx
    workbooks_sample.get_sample_master_sheet(1)

    source_file = excel_file.SourceFile("workbook_copy.xlsx", "Sheet1", "Sales Code", "Total Qty")
    destination_file = excel_file.DestinationFile("workbook_copy_1.xlsx", "Upload", "ID", "3S")

    excel_file.update(source_file, destination_file)

    modified_destination_sheet = load_workbook("workbook_copy_1.xlsx")["Upload"]

    sum_of_3s_column = 0
    for row in range(10, 532):
        value_at_3S_in_current_row = modified_destination_sheet["P" + str(row)].value

        if not value_at_3S_in_current_row:
            value_at_3S_in_current_row = 0
            
        sum_of_3s_column += value_at_3S_in_current_row
    
    assert sum_of_3s_column == 516

    



