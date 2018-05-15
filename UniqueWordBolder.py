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

    def __init__(self, UniqueWordsFileName = "uniqueWords.txt", questionFileName = "questions.txt"):
        """
        The constructor for class UniqueWordBolder.
        """

        self.qL = QuestionList(questionFileName)  # Create an object of type QuestionList
        self.uL = UniqueList(UniqueWordsFileName) # Create an object of type UniqueList

    def generateBoldedSpreadsheet(self):

        # Create the workbook
        Outfilename = "OutQuestions" + ".xlsx"
        workbook = xlsxwriter.Workbook(Outfilename)
        worksheet = workbook.add_worksheet("Questions")

        # Add a bold format object
        bold = workbook.add_format({'bold': True})
        cell_format1 = workbook.add_format({'font_size': 10, 'text_wrap': 1, 'valign': 'top', 'border': 1})

        # Bold all questions
        i = 1
        for question in self.qL.questionDatabase:
            print(question.qFields[0])
            worksheet.write_rich_string("A" + str(i), *self.boldUniqueWords(question.qFields[0], bold), cell_format1)
            worksheet.write_rich_string("B" + str(i), *self.boldUniqueWords(question.qFields[1], bold), cell_format1)
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
    uList = UniqueWordBolder() # Make an object of type UniqueWordBolder
    uList.generateBoldedSpreadsheet() # Create spreadsheet that has unique words bolded
