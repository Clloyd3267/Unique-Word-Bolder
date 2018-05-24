###################################################################################################
# Name        : UniqueWordBolder.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Class to bold unique words
###################################################################################################

import ctypes
from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

from UniqueWordBolder import *


class MainApp(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)

        #container = tk.Frame(self)
        #container.master.rowconfigure(0, weight=1)
        #container.master.columnconfigure(0, weight=1)
        #container.grid(sticky=W + E + N + S)

        Button(self, text="Question File", command=self.Question).grid(row=1, column=1, padx=50, pady=20)
        Label(self).grid(row=0, column= 1)
        Button(self, text="Unique Word File", command=self.UniqueWord).grid(row=1, column=3, padx=50, pady=20)

        Button(self, text="Bold Words", command=self.OnRun).grid(row=2, column=2, padx=50, pady=20)


    def Question(self):
        ftypes = [('Excel files', '*.xlsx'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        global ZYXQuestionFile
        ZYXQuestionFile = dlg.show()
        global ZYXColumns
        Col = simpledialog.askstring(title="Columns", prompt="Please input the columns for Q/A", initialvalue= "I,J")
        ZYXColumns = Col.split(',')

        if ZYXQuestionFile != '':
            Label(self, text='Ready!', fg='green').grid(row=0, column=1)





    def UniqueWord(self):
        ftypes = [('Excel files', '*.xlsx'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        global ZYXUniqueFile
        ZYXUniqueFile = dlg.show()

        if ZYXUniqueFile != '':
            Label(self, text="Ready!", fg='green').grid(row=0, column=3)

    def OnRun(self):

        global ZYXUniqueFile
        global ZYXQuestionFile
        global ZYXColumns

        if ZYXUniqueFile != '' and ZYXQuestionFile != '':

            uList = UniqueWordBolder(ZYXColumns, ZYXQuestionFile, ZYXUniqueFile)
            dlg = filedialog.asksaveasfilename(initialdir="/", title="Save as", initialfile = 'BoldedQuestions.xlsx',
                                     filetypes=(("Excel", "*.xlsx"), ("all files", "*.*")))
            if dlg[-5:] != ".xlsx":
                Outfilename = dlg + ".xlsx"
            else:
                Outfilename = dlg
            print(Outfilename)
            uList.generateBoldedSpreadsheet(Outfilename)
            FileOpen = messagebox.askquestion("Finished!",
                                          "You file has been bolded! \n Would you like to open it right now?")
            if FileOpen == "yes":
                ctypes.windll.ole32.CoInitialize(None)
                #upath = dlg
                pidl = ctypes.windll.shell32.ILCreateFromPathW(Outfilename)
                ctypes.windll.shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0)
                ctypes.windll.shell32.ILFree(pidl)
                ctypes.windll.ole32.CoUninitialize()

        else:
            messagebox.showerror('Error!', "Please upload files")


def main():
    global ZYXUniqueFile
    ZYXUniqueFile = ""
    global ZYXQuestionFile
    ZYXQuestionFile = ""
    app = MainApp()
    app.minsize(600, 400)
    app.mainloop()


if __name__ == "__main__":
    main()