# Author: yzjnxsantiago
# Start Date: Tuesday, November 9, 2022

#--------LIBRARIES--------#

import sys
sys.path.append('./')
from tkinter import *
import tkinter as tk
from tkinter import ttk
from zlib import Z_FIXED
from PIL import ImageTk, Image
from tkinter import filedialog
from ttkthemes import ThemedStyle
from building_blocks import *

def browse_button(file_exp, sheets, ch_vars, ch_text, file):
    
    source_sheets = []
    checkbox = []
    checkbox_vars = []
    checkbox_text = []

    filename = filedialog.askdirectory()
    
    file_exp.set(filename)
    
    if filename:
        file.configure(text=filename)
        
        source_sheets = id_sheets(filename)
        
        placement = 100

        for i in range(len(source_sheets)):
            checkbox_vars.append(IntVar())
            checkbox_text.append(StringVar())
            checkbox.append(ttk.Checkbutton(text=source_sheets[i]+ " ", variable=checkbox_vars[i]))
            checkbox[i].place(x = placement, y = 220)
            placement = placement + 78
            sheets.append(checkbox[i])
            ch_vars.append(checkbox_vars[i])
            ch_text.append(checkbox_text[i])

def browse_button_end(file_exp, file):
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("All files",
                                                        "*.*")))
    file_exp.set(filename)
    file.configure(text=filename)

def add_cells(source_entry, cell_list, all_cells, placement):

    all_cell_text = ""

    if (placement == 'source'):

        if str(source_entry.get())[0].isalpha() and str(source_entry.get())[1].isnumeric():
            all_cells.append(str(source_entry.get()))

    if (placement == 'destin'):

        if str(source_entry.get()).isalpha():
            all_cells.append(str(source_entry.get()))
        
    for i in range(len(all_cells)):
        all_cell_text = all_cell_text + all_cells[i] + "\n"
        
    all_cell_text = all_cell_text[:-1]

    cell_list.configure(text=all_cell_text)
    
    source_entry.delete(0, END)
    
    if (placement == 'source'):
        cell_list.place(x= 10, y = 275)
    
    if (placement == 'destin'):
        cell_list.place(x =350, y =275)

def nextset(checkbox, ch_vars, ch_textvars, cell_list_s, cell_list_d, all_cells, all_dcells, s_cells, mapped_sheets, d_cells):

    temp_all_cells = []
    temp_sheets = []
    temp_all_dcells = []

    for i in range(len(all_cells)):
        temp_all_cells.append(all_cells[i])

    for i in range(len(all_dcells)):
        temp_all_dcells.append(all_dcells[i])

    for i in range(len(ch_vars)):
        if (int(ch_vars[i].get()) == 1):
            checkbox[i].configure(state=DISABLED)
            temp_sheets.append(ch_textvars[i])

    mapped_sheets.append(temp_sheets)

    cell_list_s.configure(text = "")
    cell_list_d.configure(text = "")

    d_cells.append(temp_all_dcells)
    s_cells.append(temp_all_cells)

    all_cells.clear() 
    all_dcells.clear()

def finish(window):
    return window.destroy()


class setup():
    # Create the window
    def __init__(self, window):
        
        self.window = window

        # Make the window a themed window
        self.style = ThemedStyle(window)
        self.style.set_theme("clearlooks")

        # Set the background to navy blue
        window.configure(bg="light grey")

        #-------VARIABLES-------#

        self.s_path = StringVar()
        self.d_path = StringVar()
        self.all_cells = []
        self.all_dcells = []
        self.mapped_sheets = []
        self.checkbox_text = []

        #-------TITLE BAR-------#

        # Title bar image created with powerpoint
        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)


        #-------LABELS-------#

        self.source_info = ttk.Label(window,
                                text = " Input the directory for all the consistent files ")
        self.source_info.place(x=10, y = 140+20)


        self.label_file_explorer = ttk.Label(window,
                                    textvariable= self.s_path
                                    )
        self.label_file_explorer.place(x=10+100, y = 150+37)
        
                                    
        self.file_explorer = ttk.Label(window,
                                       text = "Select a Directory:")

        self.file_explorer.place(x=10+100, y = 150+37)

        self.label_file_explorer_des = ttk.Label(window,
                                    textvariable= self.d_path
                                    )
        self.label_file_explorer_des.place(x=10+100, y = 550+3)
        
        self.file_explorer_des = ttk.Label(window,
                                           text = "Select a Directory:")
        self.file_explorer_des.place(x=10+100, y = 550+3)
                                    
        

        self.select_sheets = ttk.Label(window,
                                text = "Select Sheets: ")
        self.select_sheets.place(x=10, y=220)

        self.arrow = ttk.Label(window,
                        text= "---------------->")
        self.arrow.place(x=235, y = 250)

        self.source_cells = ttk.Label(window)
        self.destination_cells = ttk.Label(window)


        #-------BUTTONS-------#

        self.source_browse = ttk.Button(text="Browse Files", 
                            command= lambda : browse_button(self.s_path, self.sheets, self.ch_vars, self.checkbox_text, self.file_explorer))
        self.source_browse.place(x= 10, y = 150+32)

        self.source_cell = ttk.Button(text="Add Cell", command= lambda: add_cells(self.source_cell, self.destination_cells, self.all_cells, "source"))
        self.source_cell.place(x = 10+130, y = 245)

        self.destination_cell = ttk.Button(text="Add Column", command= lambda: add_cells(self.destination_cell, self.source_cells, self.all_dcells, "destin"))
        self.destination_cell.place(x = 350+130, y = 245)

        self.next_set = ttk.Button(text = "Confirm Set", command= lambda: nextset(self.sheets, self.ch_vars, self.checkbox_text, self.source_cells, self.destination_cells, self.all_cells, self.all_dcells, self.s_cells, self.mapped_sheets, self.d_cells))
        self.next_set.place(x = 10, y = 600)

        self.proceed = ttk.Button(text ="Proceed", command= lambda: finish(window))
        self.proceed.place(x=900,y=600)

        self.destination_browse = ttk.Button(text="Browse Files", 
                            command= lambda : browse_button_end(self.d_path, self.file_explorer_des))
        self.destination_browse.place(x= 10, y = 550)

        #-------ENTRIES-------#

        self.source_cell = ttk.Entry()
        self.source_cell.place(x = 10, y = 250)

        self.ch_vars = []
        self.sheets = []

        self.s_cells = []
        self.d_cells = []

        self.destination_cell = ttk.Entry()
        self.destination_cell.place(x = 300+50, y = 250)

        #------CLOSE------#
        window.title('Excel to Excel')
        window.geometry("980x640+10+10")
        window.resizable(False,False)
        window.mainloop()

#-------END-------#

