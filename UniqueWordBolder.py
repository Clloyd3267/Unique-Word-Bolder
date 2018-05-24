###################################################################################################
# Name        : UniqueWordBolder.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Class to bold unique words
###################################################################################################


# External Imports
import xlsxwriter # Used to write questions out
import re # Used for pattern matching

# Project Imports
from QuestionList import *
from UniqueList import *


class UniqueWordBolder:
    """
    A class to bold unique words.

    Attributes:
        qL (QuestionList): Object of type QuestionList.
        uL (UniqueList): Object of type UniqueList.
    """

    def __init__(self, cols, questionFileName = "questions.xlsx", uniqueWordsFileName = "uniqueWords.xlsx"):
        """
        The constructor for class UniqueWordBolder.
        """

        self.qL = QuestionList(questionFileName)  # Create an object of type QuestionList
        self.uL = UniqueList(uniqueWordsFileName) # Create an object of type UniqueList
        self.cols = []

        for col in cols:
            self.cols.append(ord(col.lower()[0]) - 97)


    def generateBoldedSpreadsheet(self, Outfilename):

        # Create the workbook
        #Outfilename = "OutQuestions" + ".xlsx"
        print(Outfilename)
        workbook = xlsxwriter.Workbook(Outfilename)
        worksheet = workbook.add_worksheet("Questions")

        # Add a bold format object
        bold = workbook.add_format({'bold': True})
        cell_format1 = workbook.add_format({'font_size': 11, 'text_wrap': 1, 'valign': 'top', 'border': 1})

        # Bold all questions
        i = 0
        for question in self.qL.questionDatabase:
            j = 0
            for field in question.qFields:
                if j in self.cols:
                    worksheet.set_column(j, j, 50)
                    worksheet.write_rich_string(i, j, *self.boldUniqueWords(field, bold), cell_format1)
                else:
                    worksheet.set_column(j, j, 5)
                    worksheet.write_string(i, j, field, cell_format1)
                j += 1
            i += 1

        workbook.close() # Close workbook

    def boldUniqueWords(self, myString, boldFormat):
        """
        Function to bold unique words in a particular string.

        Parameters:
            myString (str): The input string to be bolded.
            boldFormat (xlsxwriter format object): The format to be applied to unique words.
        """

        if myString == "":
            print(myString)
        result = []
        word = ""
        match = re.search(r'According\sto.*Chapter', myString, re.IGNORECASE)
        if match:
            result.append(myString)
            return result
        for character in myString:
            if character.isalnum() or character in self.uL.partOfWord:
                word += character
            elif self.uL.isWordUnique(word):
                    result.append(boldFormat)
                    result.append(word)
                    result.append(character)
                    word = ""
            else:
                if word:
                    result.append(word)
                result.append(character)
                word = ""
        if word:
            if self.uL.isWordUnique(word):
                result.append(boldFormat)
                result.append(word)
            else:
                result.append(word)

        return result


if __name__ == "__main__":
    uList = UniqueWordBolder(["I", "J"]) # Make an object of type UniqueWordBolder
    uList.generateBoldedSpreadsheet() # Create spreadsheet that has unique words bolded
