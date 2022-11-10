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
from setup_gui import *

#-------MAIN-------#

window=tk.Tk()
setup_gui = setup(window)

source_path = str(setup_gui.s_path.get())

source_workbook_paths = find_files(".xlsx", source_path)

source_cells = setup_gui.s_cells

source_sheets = ["Project 1", "Project 2"]

destination_column = ["A", "B", "C"]

destination_workbook = xw.Book("Path to destination goes here")
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








