# Author: yzjnxsantiago
# Start Date: Tuesday, November 8, 2022

# Building Blocks: Special functions to make the program as modular as possible
# This includes functions that can obtain the cells or the modules that obtain the sheets

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

#-------LIBRARIES-------#

from ftplib import error_perm
from importlib.metadata import files
from msilib.schema import Class
from operator import concat
from urllib import response
from webbrowser import get
import xlwings as xw
from pywintypes import com_error
import os
import shutil

#-------FIND FILES-------#

def find_files(filename, search_path):
   
   '''
   Creates an array of paths to a file that contains a certain string after searching through a directory.
   
   :param str filename: A string in the filename that will be searched for. Every file that contains this string will be stored in an array.
   :param str search_path: The directory that the function will start its search from.
   :return str[] result: The array where all the paths of the files that were found are stored and returned.
   '''

   result = []

   # Walking top-down from the root
   for root, dir, files in os.walk(search_path):
      for file in files:
         if filename in file:
            result.append(os.path.join(root, file))
   return result

#-------MOVE CELLS-------#

def move_cell(row: int, cell: str, destination_column: str, sheet, destination):

    """ 
        :dependencies -- xlwings as xw

        :type row: int -- the row for which the origin cell will be copied to
        :type cell: str -- the origin cell
        :type destination_column: str -- the column for which the origin cell will be copied to
        :type sheet: Book.sheets -- the sheet where the origin cell value will be read
        :type destination: Book.sheets -- the sheet where the origin cell will be copied to

        This function reads the value from a cell in an origin excel spreadsheet and copies
        this value to another spreadsheet at a given destination excel spreadsheet
    """

    
    origin = sheet[cell].value # Store the value of the orgin cell

    destination[destination_column + str(row)].value = origin  # Moves the origin cell to the destination

#-------SHEET CHECK-------#

def id_sheets(directory: str):

    """
        :dependencies -- xlwings as xw

        :type directory: str -- the path to the parent directory where a list of excel files will be found
        :rtype sheets: List[str] -- returns all the sheet names of an excel file
        
        Goes through a directory and finds all excel files and returns a list of sheet names

    """

    books = []
    sheets = []

    books = find_files(".xlsx", directory) # Stores all the excel files in 'directory'

    # Make sure there is a spreadsheet in books
    if books[0]:
        book = xw.Book(books[0])
    
    # Append all the sheet names to 'sheets'
    for i in range(len(book.sheet_names)):
        sheets.append(book.sheet_names[i])

    return sheets

#-------PARSE EXCEL CELL RANGE-------#

def excel_range(exrange: str):

    """

        :type exrange: str -- an excel range string formatted as "XX##:XX##"
        :rtype range_array: List[str] -- a list of all the cells in the excel range

        Returns a list of all the excel cells in an excel range

    """
    
    range_array = []

    initial     = exrange[:int(len(exrange)/2)]   # Take the left side of the string up to the middle character (exlcludes the middle char)
    final       = exrange[int(len(exrange)/2)+1:] # Take the right side of the string from the middle character (exccludes the middle char)
    
    column      = find_column(initial)            

    initial_row = find_ints(initial)              
    final_row   = find_ints(final)                

    row_diference  = final_row - initial_row + 1  # Row difference adds 1 to account for the final cell

    # Stores each cell in the 'exrange' into 'range_array' 
    for i in range(row_diference):
        range_array.append(column + str(initial_row)) # The cell "COLUMN" + "ROW #" is appended
        initial_row += 1

    return range_array

def find_ints(cell: str):

    """

    :type cell: str -- an excel cell formatted as "XX##"
    :rtype ints: int -- returns the integers in order of the cell

    Takes a string and returns the integers in order of that string as a single integer

    """

    ints = str
    
    firstTime = True # Flag to initialize first time

    # Goes through a string, and makes a new string of only the ints
    for i in range(len(cell)):
        if cell[i].isnumeric():
            if firstTime:
                ints = cell[i]
                firstTime = False
            else:
                ints += cell[i]

    return int(ints) # return as an integer

def find_column(cell: str):

    """

    :type cell: str -- an excel cell formatted as "XX##"
    :rtype column: int -- returns the letters in order of the cell

    Takes a string and returns the letters in order of that string as a single string
    
    """

    column = str

    firstTime = True # Flag to initialize first time

    # Goes through a string, and makes a new string of only the ints
    for i in range(len(cell)):
        if cell[i].isalpha():
            if firstTime:
                column = cell[i]
                firstTime = False
            else:
                column += cell[i]

    return column 

def find_completed_sheets(workbook, sheet, reference: str, keyword: str, keyword_cells: str):

    
    validation_sheet = workbook.sheets[sheet]
    completed_sheets = set()

    for i in range(len(keyword_cells)):
        key_cell = validation_sheet[keyword_cells[i]].value
        if key_cell == keyword:
            completed_sheets.add(validation_sheet[reference[i]].value)
    
    return completed_sheets