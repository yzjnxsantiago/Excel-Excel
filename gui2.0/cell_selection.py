# Author: yzjnxsantiago
# Start Date: Monday, April 3, 2023

LARGEFONT =("Verdana", 35)
BACKGROUND_COLOR = "#262335"
SECONDARY_COLOR  = "#241b2f"
BUTTON_COLOR     = "#5a32fa"
BUTTON_HIGHLIGHT = "#7654ff"

import sys
sys.path.append('./')
from tkinter import *
import tkinter as tk
from tkinter import ttk
from zlib import Z_FIXED
from PIL import ImageTk, Image
from tkinter import filedialog
from building_blocks import *
#from excel_to_excel import excel_excel
import threading

def cell_selection(root: Frame):

    cells = []
    columns = []

    cell_sel_lbl = Label(root, text = "Cell Selection", font=('Calabri', 14), bg= SECONDARY_COLOR, fg='White', background=SECONDARY_COLOR)
    cell_sel_lbl.grid(row=0, column=0)

    for i in range(20):
        cells.append(Entry(root, bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR, width= 4, font=('Calabri', 14)))
        columns.append(Entry(root, bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR, width= 4, font=('Calabri', 14)))
        cells[i].grid( row =1, column=i+1, pady =4, padx=4)
        columns[i].grid(row=2, column=i+1, pady =4, padx=4)

    select_cells_btn = Button(root, text="Select", font= ('Calabri', 14), borderwidth=1, relief="ridge",     
                background= "#800020", foreground='White', activebackground="#a6022b" , activeforeground="White", cursor="hand2", state="disabled"
                )
    select_cells_btn.grid(row=1, rowspan= 2, column=0)

    new_cells_btn = Button(root, text="+", font= ('Calabri', 20), borderwidth=1, relief="ridge", width=3,   
                background= "#800020", foreground='White', activebackground="#a6022b" , activeforeground="White", cursor="hand2", state="disabled"
                )
    new_cells_btn.grid(row=1, rowspan= 2, column=21, padx=4)

    
    

