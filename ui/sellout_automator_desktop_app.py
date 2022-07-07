from tkinter import *
from tkinter import filedialog

import sys
sys.path.append('/Users/stanleykurniawan/sellout-automator')
import models.excel_file as excel_file

window = Tk()

source_file_path = ""
destination_file_path = ""

source_file = Label(window, text = "Pick Source File!")

def choose_source_file():
    global source_file_path
    source_file_path = filedialog.askopenfilename()
    source_file.config(text = source_file_path)

choose_source_file_button = Button(text="Choose Source", command = choose_source_file)

source_sheet_label = Entry()
source_sheet_label.insert(0, "Source File Sheet Label")

source_id_label = Entry()
source_id_label.insert(0, "Source ID Label")

source_data_label = Entry()
source_data_label.insert(0, "Source Data Label")

destination_file = Label(window, text = "Pick Destination File!")

def choose_destination_file():
    global destination_file_path
    destination_file_path = filedialog.askopenfilename()
    destination_file.config(text = destination_file_path)

choose_source_file_button = Button(text="Choose Destination", command = choose_destination_file)

destination_sheet_label = Entry()
destination_sheet_label.insert(0, "Source File Sheet Label")

destination_id_label = Entry()
destination_id_label.insert(0, "Source ID Label")

destination_data_label = Entry()
destination_data_label.insert(0, "Source Data Label")

source_file.pack()
choose_source_file_button.pack()
source_sheet_label.pack()
source_id_label.pack()
source_data_label.pack()

destination_file.pack()
choose_source_file_button.pack()
destination_sheet_label.pack()
destination_id_label.pack()
destination_data_label.pack()

update_status_label = Label(window, text = "Waiting...")

def update():
    source_file = excel_file.SourceFile(source_file_path, source_sheet_label.get(), source_id_label.get(), source_data_label.get())
    destination_file = excel_file.DestinationFile(destination_file_path, destination_sheet_label.get(), destination_id_label.get(), destination_data_label.get())

    excel_file.update(source_file, destination_file)
    update_status_label.config(text="Successfully Updated!")


update_destination_button = Button(text="Update!", command= update)

update_destination_button.pack()
update_status_label.pack()

window.mainloop()