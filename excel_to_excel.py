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

# Create a tkinter object
window=tk.Tk()

# Use the setup_gui for our tkinter object and return the information 
setup_gui = setup(window)

# Get the source path from the gui
source_path = str(setup_gui.s_path.get())

# Get the destination path from the gui
destination_path = str(setup_gui.d_path.get())

# Get the paths to all the excel files in the source_path folder
source_workbook_paths = find_files(".xlsx", source_path)

# Get the cells from the gui (2D Array)
source_cells = setup_gui.s_cells

# Sheets that will be used
source_sheets = ["Project 1", "Project 2"]

# An array of strings that contains the columns each cell value will go to
destination_columns = setup_gui.d_cells

# Open the destination workbook
destination_workbook = xw.Book(destination_path)
# Open the destination sheet
destination_sheet = destination_workbook.sheets["Sheet1"]

# Initialize the count for the starting row of the destination
count = 2
    
# Start by iterating through each source workbook
for workbook in source_workbook_paths:

    # If the workbook has ~$ it does not need to read it
    if "~$" in workbook:
        continue

    # Try opening the source workbook 
    try:    
        source_workbook = xw.Book(workbook)
    except:
        continue
    
    # The main algorithm to iterate though each cell that belongs to each sheet and place the values of the cells the correct location 
    # at the destination worbook
    for (sheet_cells, sheet, sheet_columns) in zip(source_cells, source_sheets, destination_columns):
        for (cell, column) in zip(sheet_cells, sheet_columns):
            move_cell(count, cell, column, source_workbook.sheets[sheet], destination_sheet)

        # Increase the row count            
        count = count + 1
    
    # Try to save and close the workbook
    try:
        source_workbook.save()
        source_workbook.close()
    except:
        print("1 Error Added")
        pass








