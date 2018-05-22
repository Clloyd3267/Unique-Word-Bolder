###################################################################################################
# Name        : UniqueList.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Class to store and manage uniques words
###################################################################################################

# External Imports
import openpyxl # For reading in unique words


class UniqueList:
    """
    A class to store UniqueList.

    Attributes:
        uniqueWords (array of str): An array to hold all of the unique words.
        partOfWord (array of str): An array to hold any characters that are not a number or letter.
    """

    uniqueWords = [] # An array to hold all of the unique words
    partOfWord = [] # An array to hold any characters that are not a number or letter.


    def __init__(self, uniqueWordsFileName = "uniqueWords.xlsx"):
        """
        The constructor for class MaterialList.

        Parameters:
            uniqueWordsFileName (str): The input filename for Unique Words, defaults to "unique.csv".
        """

        self.importUniqueWords(uniqueWordsFileName)

    def importUniqueWords(self, uniqueWordsFileName):
        """
        Function to import all unique words.

        Parameters:
            uniqueWordsFileName (str): The input filename for Unique Words.
        """

        book = openpyxl.load_workbook(uniqueWordsFileName) # Open the workbook holding the unique words

        sheet = book.active # Open the active sheet

        # Loop through all of the unique words
        for row in sheet.iter_rows(min_row = 1, min_col = 1, max_col = 1):
            uniqueWord = str(row[0].value)
            self.uniqueWords.append(uniqueWord)

            # Find all none alpha and numeric chars
            for character in uniqueWord:
                if character not in self.partOfWord and not character.isalnum() and not character.isspace():
                    self.partOfWord.append(character)

    def isWordUnique(self, testWord):
        """
        Function to import and store Unique Words.

        Parameters:
           testWord (str): The input word to be tested.

        Returns:
            True (bool): Word is unique.
            false (bool): Word is not unique.
        """

        testWord = testWord.replace(" ", "")
        testWord = testWord.lower()
        for uniqueWord in self.uniqueWords:
            uniqueWord = uniqueWord.replace(" ", "")
            uniqueWord = uniqueWord.lower()
            if testWord == uniqueWord:
                return True
        return False