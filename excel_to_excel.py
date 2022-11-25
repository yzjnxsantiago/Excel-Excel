# Author: yzjnxsantiago
# Start Date: November, 8, 2022

# Project: This script is used to take all the information from PAF funding applications
# and then organizes the information into a single paf tracker excel workbook

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

#-------LIBRARIES-------#
import sys
sys.path.append('./GUI')
from ftplib import error_perm
from importlib.metadata import files
from msilib.schema import Class
from urllib import response
import xlwings as xw
from pywintypes import com_error
from building_blocks import *
from setup_gui import *
import threading
import time

#-------MAIN-------#


if __name__ == "__main__":
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

def excel_excel(Page1: Frame, Page2: Frame, Page3: Frame, Page4: Frame):

    isRunning = [True]

    threading.Thread(target = loading, args=([Page1],[Page2],[Page3],[Page4], isRunning)).start()

    app = xw.App(visible=False)

    directory_page       = Page1
    sheet_selection_page = Page2
    cell_selection_page  = Page3
    loading_page         = Page4
    
    workbook_paths            = directory_page.get_directories()

    source_directory_path     = str(workbook_paths[0].get())
    source_workbook_paths     = find_files(".xlsx", source_directory_path)
   
    destination_path          = str(workbook_paths[1].get())

    destination_workbook      = app.books.open(destination_path)

    destination_sheet         = destination_workbook.sheets["Sheet1"]
   
    source_sheets             = ['Project 1', 'Project 2']

    cell_map                  = sheet_selection_page.get_map()

    source_cells              = []

    destination_columns       = []

    for i in range(len(cell_map)):
        source_cells.append([])
        for j in range(len(cell_map[i])):
            if cell_map[i][j]:
                source_cells[i].append(cell_map[i][j].cget('text'))
    
    for i in range(len(cell_map)):
        destination_columns.append([])
        for j in range(len(cell_map[i])):
            if cell_map[i][j]:
                destination_columns[i].append(alphabet[j])

    # Initialize the count for the starting row of the destination
    count = 2
        
    # Start by iterating through each source workbook
    for workbook in source_workbook_paths:

        # If the workbook has ~$ it does not need to read it
        if "~$" in workbook:
            continue

        # Try opening the source workbook 
        try:    
            source_workbook = app.books.open(workbook)
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
    
    isRunning[0] = False
        
def loading(Page1: Frame, Page2: Frame, Page3: Frame, Page4: Frame, isRunning: bool):

    loading_label = Page4[0].get_loading_label()

    while isRunning[0]:
        time.sleep(0.5)
        loading_label.config(text='Loading..')
        time.sleep(0.5)
        loading_label.config(text='Loading...')
        time.sleep(0.5)
        loading_label.config(text='Loading.')
    
    loading_label.config(text='Done')
    loading_label.place(x = 525, y = 700/2)



