import workbooks_sample

test_sheet = workbooks_sample.get_sample_distributor_sheet()
print(test_sheet["A1"].value)
print(test_sheet["A2"].value)
print(test_sheet["A6"].value)
