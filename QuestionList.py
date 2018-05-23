###################################################################################################
# Name        : QuestionList.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Classes to store and manage questions
###################################################################################################

# External Imports
import openpyxl # For reading in questions

class Question:
    """
    A class to store some of the attributes of a question.

    Attributes:
        qFields (array of str): Any strings that need to be bolded from question.
    """

    def __init__(self, fields):
        """
        The constructor for class Question.

        Parameters:
            fields (array of str): Any strings that need to be bolded from question.
        """

        self.qFields = fields


class QuestionList:
    """
    A class to store the questions.

    Attributes:
        questionDatabase (array of Question): The array that stores all of the questions.
    """

    questionDatabase = []  # The array that stores all of the questions

    def __init__(self, questionFileName = "questions.xlsx"):
        """
        The constructor for class Question List.

        Parameters:
            questionFileName (str): The input filename for questions, defaults to "questions.txt".
        """

        self.importQuestions(questionFileName) # Import all questions

    def importQuestions(self, questionFileName):
        """
        Function to import questions and populate QuestionList object.

        Parameters:
            questionFileName (str): The input filename for questions.
        """

        book = openpyxl.load_workbook(questionFileName)  # Open the workbook holding the unique words

        sheet = book.active  # Open the active sheet
        print("A1:", sheet['A1'].value)
        # Loop through all of the unique words
        for row in sheet.iter_rows(min_row = 1, min_col = 1):
            fields = []
            for cell in row:
                if cell.value == None:
                    fields.append("")
                else:
                    fields.append(str(cell.value))
            questionObj = Question(fields)
            self.questionDatabase.append(questionObj)

