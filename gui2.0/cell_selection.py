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

    
    

#Let me describe my project, let me know how I can organize this project, how I should divide the apps,  and if you forsee any potential issues:

#My website will first be a homepage. This homepage will showcase the legends ascension arena, the subscription plans, the elo system and account registration. This will be the first area the user will be able to access to find information about how legends ascension arena. Next once the user logs in they will have a dashboard. This dashboard will have a profile that will showcase the league of legends account (stuff like rank, level and profile name), it will also show the statistics of their matches with legends ascension arena, the teams they have been and the current team they are on as well as their schedule. Some of these things will have separate pages where the user can find even more detailed information. For example, they can click on one of the teams they were on and find out their match history. Next we will need another page or section for the user to sign up for a league. Here they will be able to sign up for a league depending on their rank, how many teams they want to play with, and how much money they want to compete for. When they sign up, this will then be tossed to the backend. The backend will need to determine which teams will need to play with each other and the schedule of the game. These games will then need to be automatically created with a tournament passcode. The game will automatically close 5 minutes after the time. After the game is completed, the teams will need to 