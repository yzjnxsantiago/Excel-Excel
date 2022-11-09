# Author: yzjnxsantiago
# Start Date: Tuesday, November 8, 2022

# Building Blocks: Special functions to make the program as modular as possible
# This includes functions that can obtain the cells or the modules that obtain the sheets


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
   for root, files in os.walk(search_path):
      for file in files:
         if filename in file:
            result.append(os.path.join(root, file))
   return result

#-------MOVE CELLS-------#

def move_cell(row, cell, destination_column, sheet, destination):

        destination[concat(destination_column, str(row))].value = sheet[cell].value

