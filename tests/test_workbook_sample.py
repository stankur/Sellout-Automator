import workbooks_sample
from openpyxl.worksheet.worksheet import Worksheet
import pytest


def test_get_distributor_description_as_id_sheet():
    test_sheet = workbooks_sample.get_sample_distributor_description_as_id_sheet()
    assert test_sheet["A2"].value == None
    assert test_sheet["A6"].value == "165/80 R13 TECHNO"
    assert test_sheet["B15"].value == 300


def test_get_distributor_code_as_id_sheet():
    test_sheet = workbooks_sample.get_sample_distributor_code_as_id_sheet()
    assert test_sheet["A1"].value == "ID"
    assert test_sheet["F1"].value == "Sales Code"
    assert test_sheet["F10"].value == "SPSR00250"


def test_get_sample_master_sheet():
    test_sheet = workbooks_sample.get_sample_master_sheet()
    assert test_sheet["A1"].value == "Form Upload Sell Out"
    assert test_sheet["B1"].value == None
    assert test_sheet["A6"].value == None
    assert test_sheet["A8"].value == "ID"
    assert test_sheet["A9"].value == None
    assert test_sheet["P8"].value == "3S"
    assert test_sheet["A85"].value == "SPSR00107"
