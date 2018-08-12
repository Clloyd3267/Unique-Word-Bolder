###################################################################################################
# Name        : UniqueWordBolderGui.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : A program to bold C&MA Bible Quizzing questions
# Github Link : https://github.com/Clloyd3267/Unique-Word-Bolder/
###################################################################################################

# External Imports
from pathlib import Path # Used for file manipulation
from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Project Imports
from UniqueWordBolder import *


class MainApp(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        self.withdraw()  # To only show the dialogs
        self.wm_iconbitmap('../Data Files/myicon.ico')

        # Input questions file dialog
        fTypes = [('Excel files', '*.xlsx')]
        dlg = filedialog.Open(title = "Choose the Input Questions File", filetypes = fTypes,
                              initialdir = Path.home() / "Documents")
        questionsFile = dlg.show()
        if questionsFile == "":
            messagebox.showerror("Error", "Error => No Input file!!!")
            return

        # Output questions save as file dialog
        exportFile = filedialog.asksaveasfilename(title = "Choose the Output File", filetypes = fTypes,
                                                  initialdir = Path.home() / "Documents", initialfile = "OutQuestions.xlsx")
        if exportFile == "":
            messagebox.showerror("Error", "Error => No Output file!!!")
            return

        # Dialog for Columns
        colDialog = simpledialog.askstring(title = "Columns", prompt = "Please input the column letters to be bolded", initialvalue = "F,G")
        if colDialog == "":
            messagebox.showerror("Error", "Error => No Column Letters !!!")
            return

        # Run the program
        uLB = UniqueWordBolder(colDialog.split(','), questionsFile, exportFile)
        messagebox.showinfo(title = "Program Complete!", message = "Questions have been bolded!")


if __name__ == "__main__":
    app = MainApp()