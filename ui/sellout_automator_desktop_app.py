from tkinter import *
from tkinter import filedialog
import custom_widgets

import sys

sys.path.append("..")
import models.excel_file as excel_file

window = Tk()

custom_widgets.UpdateForm(window).pack()

window.mainloop()
