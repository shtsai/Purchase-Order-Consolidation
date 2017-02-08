import tkinter
from tkinter import filedialog

root = tkinter.Tk()
root.withdraw()   # hide the gui form

# open up a file dialog to obtain the file path
file_path = filedialog.askopenfilename()

index = file_path.find(".xlsx")
if index == -1:
    print("invalid file path: selected file is not an .xlsx file")
else:
    new_file_path = file_path[:index] + "-統一格式" + file_path[index:]
    print(new_file_path)
