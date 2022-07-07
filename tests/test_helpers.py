import pytest
import workbooks_sample
import models.helpers as helpers


def test_is_merged_cell():
    master_sheet = workbooks_sample.get_sample_master_sheet()
    assert helpers.is_merged_cell(master_sheet, "A8")
    assert helpers.is_merged_cell(master_sheet, "A9")
    assert not helpers.is_merged_cell(master_sheet, "A10")


# Helper test
def test_find():
    master_sheet = workbooks_sample.get_sample_master_sheet()
    assert helpers.Helper(master_sheet).find("ID") == "A9"


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

    assert id_value_column.get_row_of("SLSR00103") == 10
    assert id_value_column.get_row_of("STBR00102") == 513
    assert id_value_column.get_row_of("STBR00021") == 531


def test_get_value_at():
    master_sheet = workbooks_sample.get_sample_master_sheet()
    id_value_column = helpers.Helper(master_sheet).create_value_column("ID")

    assert id_value_column.get_value_at(10) == "SLSR00103"
    assert id_value_column.get_value_at(513) == "STBR00102"
    assert id_value_column.get_value_at(531) == "STBR00021"
