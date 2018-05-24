###################################################################################################
# Name        : UniqueWordBolder.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Class to bold unique words
###################################################################################################

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
        global QuestionFile
        QuestionFile = dlg.show()
        global ZYXColumns
        Col = simpledialog.askstring(title="Columns", prompt="Please input the columns for Q/A", initialvalue= "I,J")
        ZYXColumns = Col.split(',')

        if QuestionFile != '':
            Label(self, text='Ready!', fg='green').grid(row=0, column=1)





    def UniqueWord(self):
        ftypes = [('Excel files', '*.xlsx'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        global UniqueFile
        UniqueFile = dlg.show()

        if UniqueFile != '':
            Label(self, text="Ready!", fg='green').grid(row=0, column=3)

    def OnRun(self):

        global UniqueFile
        global QuestionFile
        global ZYXColumns

        if UniqueFile != '' and QuestionFile != '':

            uList = UniqueWordBolder(ZYXColumns, QuestionFile, UniqueFile)
            dlg = filedialog.asksaveasfile(initialdir="/", title="Save as", defaultextension = '.txt',
                                     filetypes=(("Excel", "*.xlsx"), ("all files", "*.*")))
            uList.generateBoldedSpreadsheet(Outfilename = dlg.name + ".xlsx")
            messagebox.showinfo("Finished!", "Your questions have been bolded!")
        else:
            messagebox.showerror('Error!', "Please upload files")


def main():
    global UniqueFile
    UniqueFile = ""
    global QuestionFile
    QuestionFile = ""
    app = MainApp()
    app.minsize(600, 400)
    app.mainloop()


if __name__ == "__main__":
    main()