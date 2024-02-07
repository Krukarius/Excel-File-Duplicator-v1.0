import tkinter as tk
from tkinter import Tk, font
from tkinter import *
import tkinter.messagebox
from tkinter import filedialog
from tkinter.filedialog import askdirectory
import pandas as pd
import pyautogui
from pathlib import Path
from openpyxl import Workbook, load_workbook

def openFile():
    filepath = filedialog.askopenfilename(
                                          title="Please select a file",
                                          filetypes = (('Excel files', '*.xlsx'),
                                                      ('Excel macro files', '*.xlsm'))
                                          )
                                                                        
    file = open(filepath, 'r')
    wb = load_workbook(filepath, read_only=False, keep_vba=True)
    ws=wb.active
    fileonly = filepath.split('/')[-1]
    filewithoutextension = fileonly.split('-')[0]
    fileextension = fileonly.split('.')[-1]
    
    y = pyautogui.prompt(text='Please specify the number of file duplicates', title='Number of files', default='2')

    file_path = filedialog.askdirectory()
    #final_file_path = '/'.join(filepath.split('/')[:-1])

    print(y)
    for x in range(2, (int(y))+1):
            newname = "%03d" % x
            save_path = Path(file_path) / f"{filewithoutextension}- {newname}.{fileextension}"
            wb.save(save_path)
            #wb.save(f''+file_path+filewithoutextension+'- '+newname+'.'+fileextension)
            print (filewithoutextension+'- '+newname+'.'+fileextension)

rootwindow = tk.Tk()
rootwindow.geometry("500x100")
rootwindow.title("Excel file duplicator")
font.families()
windowfont=tk.font.nametofont("TkDefaultFont")
windowfont.config(
    family="Segoe Script",
    size=24,
    weight=font.BOLD)

button = tk.Button(rootwindow, text='Open Excel document',command=openFile, fg='black', bg='yellow')
button.pack()
rootwindow.mainloop()

#tkMessageBox.showinfo(title="Job complete", message="Completed! Pleasze check your destination folder!")






