from tkinter import *
from tkinter import filedialog

import sys
from tkinter import font
from typing import List, Tuple

sys.path.append("..")
import models.excel_file as excel_file


class EntryGroup:
    def __init__(self, window: Tk, label_text: str):
        self.window = window
        self.label_text = label_text

        self.label = Label(window, text=label_text, font=font.Font(weight="bold", size=14))
        self.entry = Entry(window)

    def pack(self) -> None:
        self.label.pack()
        self.entry.pack()

    def destroy(self) -> None:
        self.label.destroy()
        self.entry.destroy()

    def disable(self) -> None:
        self.entry.config(state="disabled")

    def get_entry_value(self) -> Widget:
        return self.entry.get()


class FileChooseGroup:
    def __init__(self, window: Tk, label_text: str):
        self.window = window
        self.label_text = label_text
        self.file_destination = ""

        self.label = Label(window, text=label_text, font=font.Font(weight="bold", size=14))
        self.chosen_file_label = Label(window, text="No File Chosen Yet!", bg="#7F7F7F")
        self.choose_file_button = Button(
            window, text="Choose File", command=self._choose_file
        )

    def _choose_file(self) -> None:
        source_file_path = filedialog.askopenfilename()
        self.chosen_file_label.config(text=source_file_path)

    def pack(self) -> None:
        self.label.pack()
        self.chosen_file_label.pack()
        self.choose_file_button.pack()

    def destroy(self) -> None:
        self.label.destroy()
        self.chosen_file_label.destroy()
        self.choose_file_button.destroy()

    def disable(self) -> None:
        self.choose_file_button.config(state="disabled")

    def get_chosen_file_path(self) -> str:
        self.chosen_file_label.cget("text")


class UpdateForm:
    def __init__(self, window: Tk):
        self.source_file_choose_group = FileChooseGroup(window, "Choose Source File")
        self.source_sheet_label = EntryGroup(
            window, "Choose Source File Sheet Label"
        )
        self.source_id_label = EntryGroup(
            window, "Choose Source File ID Label"
        )
        self.source_data_label = EntryGroup(
            window, "Choose Source File Data Label"
        )

        self.destination_file_choose_group = FileChooseGroup(
            window, "Choose Destination File"
        )
        self.destination_sheet_label = EntryGroup(
            window, "Choose Destination File Sheet Label"
        )
        self.destination_id_label = EntryGroup(
            window, "Choose Destination File ID Label"
        )
        self.destination_source_data_label = EntryGroup(
            window, "Choose Destination File Data Label"
        )

    def pack(self) -> None:
        self.source_file_choose_group.pack()
        self.source_sheet_label.pack()
        self.source_id_label.pack()
        self.source_data_label.pack()

        self.destination_file_choose_group.pack()
        self.destination_sheet_label.pack()
        self.destination_id_label.pack()
        self.destination_source_data_label.pack()

    def destroy(self) -> None:
        self.source_file_choose_group.destroy()
        self.source_sheet_label.destroy()
        self.source_id_label.destroy()
        self.source_data_label.destroy()

        self.destination_file_choose_group.destroy()
        self.destination_sheet_label.destroy()
        self.destination_id_label.destroy()
        self.destination_source_data_label.destroy()

    def disable(self) -> None:
        self.source_file_choose_group.disable()
        self.source_sheet_label.disable()
        self.source_id_label.disable()
        self.source_data_label.disable()

        self.destination_file_choose_group.disable()
        self.destination_sheet_label.disable()
        self.destination_id_label.disable()
        self.destination_source_data_label.disable()

    def get_info(self) -> Tuple[Tuple[str, str, str, str], Tuple[str, str, str, str]]:
        return (
            (
                self.source_file_choose_group.get_chosen_file_path(),
                self.source_sheet_label.get_entry_value(),
                self.source_id_label.get_entry_value(),
                self.source_data_label.get_entry_value(),
            ),
            (
                self.destination_file_choose_group.get_chosen_file_path(),
                self.destination_sheet_label.get_entry_value(),
                self.destination_id_label.get_entry_value(),
                self.destination_source_data_label.get_entry_value(),
            ),
        )
