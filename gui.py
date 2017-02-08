import tkinter
from tkinter import filedialog

root = tkinter.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)

