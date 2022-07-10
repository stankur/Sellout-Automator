from tkinter import *
from tkinter import filedialog
import custom_widgets

import sys

from ui.custom_widgets import UpdateForm

sys.path.append("..")
import models.excel_file as excel_file

window = Tk()

def run_update(source_file_info, destination_file_info):
    source_file = excel_file.SourceFile(*source_file_info)
    destination_file = excel_file.DestinationFile(*destination_file_info)

    excel_file.update(source_file, destination_file)

def restart():
    global current_update_form
    global state_change_button
    global state_info_label

    current_update_form.enable()
    current_update_form.empty()
    state_change_button.config(text="Update", command = request_update)
    state_info_label.config(text="")


def request_update():
    global current_update_form
    global state_change_button
    global state_info_label

    update_info = current_update_form.get_info()
    source_file_info = update_info[0]
    destination_file_info = update_info[1]

    try:
        state_info_label.config(text="")

        run_update(source_file_info, destination_file_info)

        current_update_form.disable()
        state_change_button.config(text="Restart", command= restart)
        state_info_label.config(text="Succesfully Updated! Let's go again?", fg="green")

    except Exception as error:
        state_info_label.config(text=str(error), fg="red")


current_update_form: UpdateForm = custom_widgets.UpdateForm(window)
state_change_button = Button(window, text="Update", command = request_update)
state_info_label = Label(window, text="", fg="red")

current_update_form.pack()
state_change_button.pack()
state_info_label.pack()


window.mainloop()
