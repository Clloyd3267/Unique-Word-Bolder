###################################################################################################
# Name        : QuestionList.py
# Author(s)   : Chris Lloyd, Andrew Southwick
# Description : Classes to store and manage questions
###################################################################################################


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

    def __init__(self, questionFileName = "questions.txt"):
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

        questionFile = open(questionFileName, "r", encoding = 'UTF-8')

        # For each line in file, place into proper question object
        for question in questionFile:
            question = question.rstrip()
            fields = question.split("$")
            if len(fields) < 2:
                print("Error!!! => empty field!")
                continue
            questionObj = Question(fields[:2])
            self.questionDatabase.append(questionObj)
        questionFile.close()
