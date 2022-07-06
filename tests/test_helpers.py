import pytest
import workbooks_sample
import models.helpers as helpers

# Helper test
def test_get_column():
    master_sheet = workbooks_sample.get_sample_master_sheet()
    assert helpers.Helper(master_sheet).get_column("3S") == "P"
    assert helpers.Helper(master_sheet).get_column("STS MU") == "Q"


def test_create_value_column():
    distributor_sheet = workbooks_sample.get_sample_distributor_code_as_id_sheet()
    id_value_column = helpers.Helper(distributor_sheet).create_value_column(
        "Sales Code"
    )

    assert id_value_column.column == "F"
    assert id_value_column.get_start_cell_row() == 2
    assert id_value_column.get_end_cell_row() == 28

    try:
        helpers.Helper(distributor_sheet).create_value_column("Upload Answer")

        pytest.fail("expeccted empty value column to throw exception")
    except:
        Exception

    master_sheet = workbooks_sample.get_sample_master_sheet()
    id_value_column = helpers.Helper(master_sheet).create_value_column("ID")

    assert id_value_column.column == "A"
    assert id_value_column.get_start_cell_row() == 10
    assert id_value_column.get_end_cell_row() == 531


# ValueColumn test
def test_get_row_of():
    master_sheet = workbooks_sample.get_sample_master_sheet()
    id_value_column = helpers.Helper(master_sheet).create_value_column("ID")

    id_value_column.get_row_of()
