import tkinter
from tkinter import filedialog

def get_file_path():
    root = tkinter.Tk()
    root.withdraw()   # hide the gui form

    # open up a file dialog to obtain the file path
    file_path = filedialog.askopenfilename()

    index = file_path.find(".xlsx")
    if index == -1:
        print("invalid file path: selected file is not an .xlsx file")
        return -1
    else:
        return file_path


def generate_new_filename(file_path, s):
    index = file_path.find(".xlsx")
    if index == -1:
        print("invalid file path: selected file is not an .xlsx file")
        return -1
    else:
        new_file_path = file_path[:index] + "-" + s + file_path[index:]
        return new_file_path


