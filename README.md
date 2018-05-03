# Unique-Word-Bolder
A set of Python files to bold unique words

### Getting Started

#### Prerequisites
This project is built using Python 3.6 and the xlxswriter library. make sure you have them installed and working from the links below:

* (https://www.python.org/downloads/release/python-365/)
* (http://xlsxwriter.readthedocs.io/)

#### Input files
The two files that are used to import data are:

* ##### uniqueWords.txt
A text file holding one string per line corresponding to a single unique word. 
There must be no blank lines and no spaces!
Encoding must be in [UTF-8](https://en.wikipedia.org/wiki/UTF-8).<br/>Example: UniqueWord.

* ##### questions.txt
A text file holding exactly two strings per line delimited by the "$" character. These strings correspond to the text where unique words need to be bolded. There must be no blank lines!
Encoding must be in [UTF-8](https://en.wikipedia.org/wiki/UTF-8).<br/>Example: Question$Answer.

#### Running the program
Once the two input files are populated, run the program by running UniqueWordBolder.py.

The bolded result should be found in outQuestions.xlsx. 

