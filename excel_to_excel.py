# Author: yzjnxsantiago
# Start Date: November, 8, 2022

# Project: This script is used to take all the information from PAF funding applications
# and then organizes the information into a single paf tracker excel workbook

#-------LIBRARIES-------#

from ftplib import error_perm
from importlib.metadata import files
from msilib.schema import Class
from operator import concat
from urllib import response
import xlwings as xw
from pywintypes import com_error
from building_blocks import *

#-------MAIN-------#

# setup()

source_workbook_paths = find_files(".xlsx", "C:/Users/ssira/Documents/Python Scripts/Excel-Excel/Source Workbooks Testing/")

source_cells = [["A1", "B1"], ["A2", "B2"]]
source_sheets = ["Project 1", "Project 2"]

destination_column = ["A", "B", "C"]

destination_workbook = xw.Book("C:/Users/ssira/Documents/Python Scripts/Excel-Excel/Destination Workbook Testing/Destination.xlsx")
destination_sheet = destination_workbook.sheets["Sheet1"]

count = 1

for workbook in source_workbook_paths:

    if "~$" in workbook:
        continue

    source_workbook = xw.Book(workbook)
    
    for (sheet_cells, sheet) in zip(source_cells, source_sheets):
            for (cell, column) in zip(sheet_cells, destination_column):
                move_cell(count, cell, column, source_workbook.sheets[sheet], destination_sheet)
                
            count = count + 1








