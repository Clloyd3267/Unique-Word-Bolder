###################################################################################################
# Name        : UniqueWordBolder.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Class to bold unique words
###################################################################################################

from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox

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
        ftypes = [('Text files', '*.txt'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        global QuestionFile
        QuestionFile = dlg.show()

        if QuestionFile != '':
            Label(self, text=QuestionFile, fg='green').grid(row=0, column=1)




    def UniqueWord(self):
        ftypes = [('Text files', '*.txt'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes=ftypes)
        global UniqueFile
        UniqueFile = dlg.show()

        if UniqueFile != '':
            Label(self, text=UniqueFile, fg='red').grid(row=0, column=3)

    def OnRun(self):

        global UniqueFile
        global QuestionFile

        if UniqueFile != '' and QuestionFile != '':
            print(UniqueFile, QuestionFile)

            uList = UniqueWordBolder(UniqueFile, QuestionFile)
            print("It works, I think...)")
            uList.generateBoldedSpreadsheet()
            print("Does It Still Work?")
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