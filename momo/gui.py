import tkinter
from tkinter import filedialog
import os

def get_file_path():
    root = tkinter.Tk()
    root.withdraw()   # hide the gui form

    # open up a file dialog to obtain the file path
    file_path = filedialog.askopenfilename()

    index = file_path.find(".xls")
    if index == -1:
        print("invalid file path: selected file is not an .xls file")
        return -1
    else:
        return file_path


def generate_new_filename(file_path, s):
    # need to determine platform, because directory representation is different
    if os.name == "nt":
        slash = '\\'
    else:
        slash = "/"

    index = file_path.find(".xls")
    if index == -1:
        print("invalid file path: selected file is not an .xls file")
        return -1
    else:
        new_file_path = file_path[:index] + "-" + s + ".xlsx"
        # get filename
        filename = os.path.basename(new_file_path)
        # change path to current working directory
        new_file_path = os.getcwd() + slash + filename

        return new_file_path

