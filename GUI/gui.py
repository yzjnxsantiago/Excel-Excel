# Author: yzjnxsantiago
# Start Date: Tuesday, November 9, 2022

__name__ = "__main__"

LARGEFONT =("Verdana", 35)
BACKGROUND_COLOR = "#262335"
SECONDARY_COLOR  = "#241b2f"
BUTTON_COLOR     = "#5a32fa"
BUTTON_HIGHLIGHT = "#7654ff"

#--------LIBRARIES--------#
import sys
sys.path.append('./')
from tkinter import *
import tkinter as tk
from tkinter import ttk
from zlib import Z_FIXED
from PIL import ImageTk, Image
from tkinter import filedialog
from building_blocks import *
from excel_to_excel import excel_excel
import threading

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']

def browse_button(directory_label: Label, frame: Frame, sheet_frame: Frame, controller):
    
    """_summary_ 
    This function is intended to let the user select a directory so that has
    simillarly formatted excel files

    """

    files_list = [] 
    source_sheets = []
    sheet_variable = controller.get_frame(Page1).get_sheet_variables()[0]
    text_sheet_variable = controller.get_frame(Page1).get_sheet_variables()[1]
    sheet_graphics = controller.get_frame(Page1).get_sheet_variables()[2]

    # Initial Placements

    x1 = 275
    x2 = 275
    y1 = 160 + 85+20
    y2 = y1 + 40

    placement_x = 290
    placement_y = 150 + 80

    
    # Prompt the user for a directory 
    filename = filedialog.askdirectory()
    directory_label.set(filename)

    files_list = find_files(".xlsx", filename) # List of excel files in the 'filename' directory 

    # Iterate through the file list, create a label for each file in the list and shift the positioning by 25 pixels down
    for i in range(len(files_list)):
        listBox(frame, x1,y1,x2,y2,files_list[i])
        y1 += 25
        y2 += 25

    source_sheets = id_sheets(filename) # List of sheet names in the first file of the directory

    # Create a checkbox for each sheet in the excel file
    for i in range(len(source_sheets)):
            sheet_variable.append(IntVar()) # Initializing checkbox check var
            text_sheet_variable.append(StringVar()) # Initializing checkbox text var
            # Create a checkbox object and store in an array
            sheet_graphics.append(Checkbutton(sheet_frame, variable=sheet_variable[i], textvariable= text_sheet_variable[i], font= ('Calabri', 10),
                                  background=BUTTON_HIGHLIGHT, foreground='white',  borderwidth = 1, relief="ridge",  height=2,
                                  activebackground=BUTTON_COLOR, activeforeground='White', selectcolor= BACKGROUND_COLOR )) 
            text_sheet_variable[i].set(source_sheets[i]) # Set the text variable
            sheet_graphics[i].place(x = placement_x, y = placement_y) # Place the checkbox
            sheet_graphics[i].update() # Update the gui so width can be obtained
            sheet_graphics_width = sheet_graphics[i].winfo_width() 
            placement_x = placement_x +  sheet_graphics_width + 15 # Place the checkbox object relative to the previous checkbox
            
            # Horizontal limit for the checkboxes before shifting it down
            if placement_x > 1000: 
                placement_x = 290
                placement_y += 50

def browse_button_end(directory_label):
    # Get the excel file for the destination 
    filename = filedialog.askopenfilename(title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("All files",
                                                        "*.*")))
    directory_label.set(filename)

class tkinterApp(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
         
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        self.geometry("1200x700")
        self.resizable("False", "False")
        self.title("Excel to Excel")
        
         
        # creating a container
        self.container = tk.Frame(self) 
        self.container.pack(side = "top", fill = "both", expand = True)
  
        self.container.grid_rowconfigure(0, weight = 1)
        self.container.grid_columnconfigure(0, weight = 1)
  
        # initializing frames to an empty array
        self.frames = {} 

        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, Page1, Page2, Page3, Page4, SheetValidation):
            
            
            frame = F(self.container, self)
  
            # initializing frame of that object from
            # startpage, page1, page2 respectively with
            # for loop
            self.frames[F] = frame
  
            frame.grid(row = 0, column = 0, sticky ="nsew")
            
  
        self.show_frame(StartPage)  
  
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        x = self.frames
        frame = self.frames[cont]
        frame.tkraise()

    def get_frame(self, cont):
        return self.frames[cont]

    def add_frame(self, new_frame):
        frames_len = len(self.frames)
        new_frame.grid(row = 0, column = 0, sticky ="nsew")
        self.frames["frame" + str(frames_len)] = new_frame
        pass
    
    def get_container(self):
        x = self.container , self
        return self.container , self

    def get_frames_len(self):
        return len(self.frames)

class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        
        tk.Frame.__init__(self, parent)

        #-------BACKGROUND SETUP-------#

        # Set the background color  
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR, borderwidth = 0, highlightthickness = 0)
        style.configure('TLabelframe.Label', font =('Arial', 15))
        style.configure('TLabelframe.Label', foreground = "Light Grey")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        
        #------TITLE BAR------#

        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(self, image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)
        # Button for creating a new page
        create = Button(self, text ="+", font=('Calibri', 25), fg= "white", bg="#5615DE", activebackground='#6017F9',
                            activeforeground='white', width = 4,
                            command = lambda : controller.show_frame(Page1))
        create.place(x = 525, y = 700/2)

class Page1(tk.Frame):
     
    def __init__(self, parent, controller):
        
        tk.Frame.__init__(self, parent)
        
        #-------BACKGROUND SETUP-------#

        # Set the background color  
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR)
        style.configure('TLabelframe.Label', font =('Calibri', 15, 'bold'))
        style.configure('TLabelframe.Label', foreground = "White")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        style.configure('TLabel', font = ('Calabri', 12))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)

        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(self, image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)
        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=475, width=250)
        self.label.update()
        controller.show_frame(StartPage)
        menu_bar.place(x = 10, y = self.label.winfo_height() + 10) 
        
        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y = self.label.winfo_height() + 40)

        nav_sel_sheet = Button(self, text="Sheet Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White", width= 20, 
                            command = lambda: controller.show_frame(Page2))
        nav_sel_sheet.place(x = 20 , y = self.label.winfo_height() + 80)      
      
        #-----BROWSE FRAME-----#
        
        browse_frame = ttk.LabelFrame(self, text='Source', height = 100, width = 900)
        browse_frame.place(x = 275, y = self.label.winfo_height() + 10)

        #-Buttons-#

        browse_source  = ttk.Button(self, text="Browse Files", 
                            command= lambda : browse_button(self.directorystr, self, controller.get_frame(Page2), controller))
        browse_source.place(x = 290, y = self.label.winfo_height() + 35 )
        
        #-Labels-#

        self.directorystr = StringVar()
        source_directory = ttk.Label(self, textvariable=self.directorystr)
        source_directory.place(x = 420, y = self.label.winfo_height() + 40)

        #------MAIN SECTION------#
        #-Buttons-#

        next = ttk.Button(self, text='Next', command = lambda : controller.show_frame(Page2))
        next.place(x = 1065, y = 650)

        #-Labels-#

        files = ttk.Label(self, text=' Source Files ', font=('Calabri', 12), borderwidth=2, relief="groove")
        files.place(x=275, y =self.label.winfo_height() + 115)

        #-----DESTINATION SECTION-----#

        destination_frame = ttk.LabelFrame(self, text='Destination', height = 100, width = 900)
        destination_frame.place(x = 275, y = 535)

        #-Buttons-#

        browse_destination  = ttk.Button(self, text="Browse Files", 
                            command= lambda : browse_button_end(self.directory_des))
        browse_destination.place(x = 290, y = 570)
       
        self.directory_des = StringVar()
        destination_directory = ttk.Label(self, textvariable=self.directory_des, font = ('Calabri', 12))
        destination_directory.place(x = 420, y = 572)

        #----------VARIABLES-----------#
        
        self.sheet_variable = []
        self.text_sheet_variable = []
        self.sheet_checkboxes = []

    def get_sheet_variables(self):
        return [self.sheet_variable, self.text_sheet_variable, self.sheet_checkboxes]

    def get_directories(self):
        return [self.directorystr, self.directory_des]

def listBox(frame, x1, y1, x2, y2, filename):

    # Limit for the listBox
    if y2 > 605/1.25:
        return
 
    label = ttk.Label(frame, text=filename, font= ('Calabri', 10), borderwidth=2, relief="solid", width= 128) # Create a label
    label.place(x = x2, y = y2)
  
class Page2(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR)
        style.configure('TLabelframe.Label', font =('Calibri', 15, 'bold'))
        style.configure('TLabelframe.Label', foreground = "White")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        style.configure('TLabel', font = ('Calabri', 15))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)

        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(self, image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)
        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=475, width=250)
        self.label.update()
        controller.show_frame(StartPage)
        menu_bar.place(x = 10, y = self.label.winfo_height() + 10)

        nav_sel_sheet = Button(self, text="Sheet Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White", width= 20, 
                            command = lambda: controller.show_frame(Page2))
        nav_sel_sheet.place(x = 20 , y = self.label.winfo_height() + 80) 
        
        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y = self.label.winfo_height() + 40)
      

        #-------SHEET SELECTION-----#

        sheet_selection = ttk.LabelFrame(self, text='Sheet Selection', height = 250, width = 900)
        sheet_selection.place(x = 275, y = self.label.winfo_height() + 20)

        #------MAIN SECTION-------#

        type_a = Button(self, text="Sheet Validation", font= ('Calabri', 15), borderwidth=1, relief="ridge", width= 15, height = 5, 
                        background= BUTTON_HIGHLIGHT, foreground='White', activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2",
                        command = lambda: controller.show_frame(SheetValidation)
                        )
        type_a.place(x=275, y = self.label.winfo_height() + 300)

        type_b = Button(self, text="Cell to Column", font= ('Calabri', 15), borderwidth=1, relief="ridge", width= 15, height = 5, 
                        background= BUTTON_HIGHLIGHT, foreground='White', activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2",
                        command= lambda: [controller.add_frame(Page3(controller.get_container()[0], controller.get_container()[1])), 
                        create_columns(controller.get_frame("frame" + str(controller.get_frames_len() - 1)), self.columns, self.cell_map, controller),
                        controller.show_frame("frame" + str(controller.get_frames_len() - 1))] 
                        )
        type_b.place(x=275+200, y = self.label.winfo_height() + 300)

        finish = Button(self, text="Finish", font= ('Calabri', 17), borderwidth=1, relief="ridge", width=8,       
                        background= "#800020", foreground='White', activebackground="#a6022b" , activeforeground="White", cursor="hand2",
                        command = lambda: [controller.show_frame(Page4), threading.Thread(target = excel_excel, args=[controller.get_frame(Page1), controller.get_frame(Page2),
                                                                                                                       controller.get_frame(Page3), controller.get_frame(Page4),
                                                                                                                      controller.get_frame(SheetValidation)]).start()
                                                                                     ]
                        )
        finish.place(x = 1070, y = 645)

        self.columns = []
        self.click_count = []
        self.cell_map = []
        self.sheet_map = []
        self.disabled_checkboxes = set()

    def get_columns(self):
        return self.columns

    def get_map(self):
        return self.cell_map

    def get_sheet_map(self):
        return self.sheet_map

    def get_click_count(self):
        return self.click_count

    def get_disabled_checkboxes(self):
        return self.disabled_checkboxes

class Page3(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR)
        style.configure('TLabelframe.Label', font =('Calibri', 15, 'bold'))
        style.configure('TLabelframe.Label', foreground = "White")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        style.configure('TLabel', font = ('Calabri', 15))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)
        
        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(self, image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)
        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=475, width=250)
        self.label.update()
        menu_bar.place(x = 10, y = self.label.winfo_height() + 10) 
        
        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y = self.label.winfo_height() + 40)

        nav_sel_sheet = Button(self, text="Sheet Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White", width= 20, 
                            command = lambda: controller.show_frame(Page2))
        nav_sel_sheet.place(x = 20 , y = self.label.winfo_height() + 80)

        #-------CELL SELECTION-------#

        sheet_selection = ttk.LabelFrame(self, text='Cell Selection', height = 90, width = 170)
        sheet_selection.place(x = 275, y = self.label.winfo_height() + 20)

        source_cell = Entry(self, bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR, width= 5)
        source_cell.place(x = 290, y = self.label.winfo_height() + 60)

        cell_count = [0]

        confirm_cell = Button(self,text="Confirm Cell", borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                              activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2", command= lambda: add_cells(self, self.cells, cell_count, controller, source_cell)
                              )
        confirm_cell.place(x= 336, y = self.label.winfo_height() + 58)

        #------CELL------#

        self.cells = []

        cell_frame = ttk.LabelFrame(self, text='Cell', height = 90, width = 90)
        cell_frame.place(x = 450, y = self.label.winfo_height() + 20)

        #-----FINISH-----#

        finish = Button(self, text="Next Sheets", font= ('Calabri', 17), borderwidth=1, relief="ridge",     
                        background= "#800020", foreground='White', activebackground="#a6022b" , activeforeground="White", cursor="hand2",
                        command= lambda: [controller.show_frame(Page2) , '''clear_page(self, controller)''']
                        )
        finish.place(x = 1045, y = 645)

#---------Page3 FUNCTIONS----------#

def create_columns(frame, columns, cell_map, controller):

    sheet_map = controller.get_frame(Page2).get_sheet_map() # 2D array storing the sheets used for a cell creation
    sheet_variables = controller.get_frame(Page1).get_sheet_variables()[0] # Sheet variables are needed to disable them
    text_sheet_variable =  controller.get_frame(Page1).get_sheet_variables()[1] 
    sheet_checkboxes = controller.get_frame(Page1).get_sheet_variables()[2] # Sheet checkboxes needed to disable them
    disabled_checkboxes = controller.get_frame(Page2).get_disabled_checkboxes() # List of disabled checkboxes

    cell_map.append([]) # Everytime a new Page3 is loaded another array has to be made within the 2D cell_map
    sheet_map.append([]) # Same for sheet map
    columns.append([]) # Same for columns

    column_x = 350 # Initial x position for the columns
    column_y = 300 # Initial y position for the columns

    column_x_start = column_x
    column_y_start = column_y

    click_count = controller.get_frame(Page2).get_click_count() # To check how many times the arrow keys have been clicked. Right arrow is +1 left is -1

    i = 0 

    while column_x < 1000: # Dont make any new columns after this position
        columns[len(cell_map)-1].append(ttk.LabelFrame(frame, text=alphabet[i], height = 90, width = 90)) # Create a new label within the current cell map
        columns[len(cell_map)-1][i].place(x = column_x, y = column_y) # Place the column
        columns[len(cell_map)-1][i].update() # Update the column to get info on width
        column_x += columns[len(cell_map)-1][i].winfo_width() + 10 # Increase the x position 
        cell_map[len(cell_map)-1].append(None) # Initialize one empty column in the cell map
        i += 1 
    
    for i in range(len(sheet_variables)):  # Iterate through sheets
        sheetisChecked = int(sheet_variables[i].get()) # Get the sheet
        if sheetisChecked == 1 and not str(text_sheet_variable[i].get()) in disabled_checkboxes: # If the sheet is checked and is not disabled
            sheet_map[len(cell_map)-1].append(str(text_sheet_variable[i].get())) # Store the sheets in the sheet map
            disabled_checkboxes.add(str(text_sheet_variable[i].get())) # Add the checkboxes to the list of disabled_checkboxes
            sheet_checkboxes[i].configure(state='disabled') #Disable the checkbox
            
    
    column_x1_last = column_x # Store the last column x position
    column_y1_last = column_y # Store the last column y position

    click_count.append(0) 

    right_button = Button(frame, text=">", font =('Calabri', 15), borderwidth=1, relief="groove", background=BUTTON_HIGHLIGHT, foreground="White", 
                      activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2", command = lambda: shift_columns_left(frame, columns, click_count, left_button,
                      controller)) # Create the right click button
    right_button.place(x = column_x1_last + 10, y = column_y1_last + 25)

    left_button = Button(frame, text="<", font =('Calabri', 15), borderwidth=1, relief="groove", background=BUTTON_HIGHLIGHT, foreground="White", 
                      activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2", command = lambda: shift_columns_right(frame, columns, click_count, left_button,
                      controller),
                      state='disabled') # Left click button
    left_button.place(x = column_x_start - 40, y = column_y_start + 25)

    return columns

def shift_columns_left(frame, columns, click_count, left_button, controller):

    cell_map = controller.get_frame(Page2).get_map() # Get the cell map
    columns[len(cell_map)-1][click_count[0]].pack_forget() # Remove the right most column
    first_column_x1 = columns[len(cell_map)-1][0].winfo_x() # x position of the first column
    first_column_y1 = columns[len(cell_map)-1][0].winfo_y() # y position of the first column

    first_loop = True # First loop flag

    # This code shifts columns and any cells that are in the columns to the left
    for i in range(click_count[0], len(columns[len(cell_map)-1])): # Iterate through the columns
        if i + 1 < len(columns[len(cell_map)-1]): 
            columns[len(cell_map)-1][i+1].place(x = first_column_x1, y = first_column_y1)
            if cell_map[len(cell_map)-1][i+1]:
                if i + 1 != click_count[0]:
                    cell_map[len(cell_map)-1][i+1].place(x = first_column_x1 + 30, y = first_column_y1 + 30)
            elif first_loop and cell_map[len(cell_map)-1][i]:
                    cell_map[len(cell_map)-1][i].lower()
            first_column_x1 += columns[len(cell_map)-1][i].winfo_width() + 10
        else:
            columns[len(cell_map)-1].append(ttk.LabelFrame(frame, text=alphabet[i+1], height = 90, width = 90))
            cell_map[len(cell_map)-1].append(None)      
            columns[len(cell_map)-1][i+1].place(x = first_column_x1, y = first_column_y1)
            if cell_map[len(cell_map)-1][i+1]:
                cell_map[len(cell_map)-1][i+1].lift()
        
        first_loop = False

    click_count[0] += 1 # The click_count has moved once

    left_button.configure(state="active") # Enable the left button since it is not longer 

def shift_columns_right(frame, columns, click_count, left_button, controller):

    cell_map = controller.get_frame(Page2).get_map()

    columns_length = len(columns[len(cell_map)-1])

    last_column_x1 = columns[len(cell_map)-1][columns_length-1].winfo_x()
    last_column_y1 = columns[len(cell_map)-1][columns_length-1].winfo_y()

    rightmost_column = columns[len(cell_map)-1][len(columns[len(cell_map)-1])-1]
    rightmost_column.destroy()
    
    columns[len(cell_map)-1].pop()

    columns_length = len(columns[len(cell_map)-1])

    first_loop = True
    
    for i in range(columns_length):
        if columns_length - i - 1 > click_count[0] - 1:
            columns[len(cell_map)-1][columns_length - i - 1].place(x = last_column_x1, y = last_column_y1)
            if cell_map[len(cell_map)-1][columns_length - i -1]:
                cell_map[len(cell_map)-1][columns_length - i - 1].place(x = last_column_x1 + 30, y = last_column_y1 + 30)
            elif first_loop and cell_map[len(cell_map)-1][columns_length- i]:
                cell_map[len(cell_map)-1][columns_length-i].lower()
            last_column_x1 -= columns[len(cell_map)-1][columns_length - i- 1].winfo_width() + 10
        elif columns_length - i - 1 == click_count[0] - 1:
            if cell_map[len(cell_map)-1][columns_length - i - 1]:
                 cell_map[len(cell_map)-1][columns_length- i - 1].lift()
        else:
            break

        first_loop = False
        
            

    click_count[0] -= 1

    if click_count[0] == 0:
        left_button.configure(state="disabled")

def add_cells(frame, cells, cell_count, controller, cell_entry):

    cells.append(Button(frame, text=cell_entry.get(), font =('Calabri', 15), borderwidth=1, relief="groove", background=BUTTON_HIGHLIGHT, foreground="White", 
                        activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2"))

    cell_map = controller.get_frame(Page2).get_map()
    click_count = controller.get_frame(Page2).get_click_count()

    current_cell = cells[cell_count[0]]
    
    current_cell.place(x=475, y= 205)

    current_cell.lift()

    cell_frame_index = len(cell_map)-1

    def drag_start(event):
        widget = event.widget
        widget.startX = event.x
        widget.startY = event.y

    def drag_motion(event):
        widget = event.widget
        x = widget.winfo_x() - current_cell.startX + event.x
        y = widget.winfo_y() - current_cell.startY + event.y
        widget.place(x=x,y=y)
        widget.lift()
    
    def clip_to_destination(event):
        widget = event.widget
        x = widget.winfo_x()
        y = widget.winfo_y()
        frame = controller.get_frame(Page2)
        columns = frame.get_columns()
        
        for i in range(click_count[0], len(columns[cell_frame_index])):
            column_x1 = columns[cell_frame_index][i].winfo_x()
            column_x2 = column_x1 + columns[cell_frame_index][i].winfo_width()
            column_y1 = columns[cell_frame_index][i].winfo_y()
            column_y2 = column_y1 + columns[cell_frame_index][i].winfo_height()
            if x > column_x1 and x < column_x2 and y > column_y1 and y < column_y2:
                widget.place(x=column_x1+ ((column_x2-column_x1)/3.25) , y = column_y1 + ((column_y2-column_y1)/3.25))
                for j in range(len(cell_map[len(cell_map)-1])):
                    if current_cell == cell_map[len(cell_map)-1][j]:
                        break
                cell_map[len(cell_map)-1][j] = None
                cell_map[len(cell_map)-1][i] = current_cell
                break
        
    current_cell.bind('<Button-1>', drag_start)  
    current_cell.bind('<B1-Motion>', drag_motion)
    current_cell.bind('<ButtonRelease>', clip_to_destination)
    
    cell_entry.delete(0, 'end')
    cell_count[0] += 1

def clear_page(frame: Frame, controller):

    cell_map = controller.get_frame(Page2).get_map()
    columns  = controller.get_frame(Page2).get_columns()

    for i in range(len(cell_map[len(cell_map)-1])):
        if cell_map[len(cell_map)-1][i]: 
            cell_map[len(cell_map)-1][i].place(x=None, y=None)
            cell_map[len(cell_map)-1][i].update()
    
    initial_column_length = len(columns)

    for i in range(initial_column_length):
        end_array = initial_column_length - i - 1
        columns[end_array].destroy()
        columns.pop()

#----------------------------------#

class SheetValidation(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR)
        style.configure('TLabelframe.Label', font =('Calibri', 15, 'bold'))
        style.configure('TLabelframe.Label', foreground = "White")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        style.configure('TLabel', font = ('Calabri', 15))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)

        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=698, width=250) 
        menu_bar.place(x = 10, y =0)
        
        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection", borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y =30)
        
        self.keyword_text = StringVar()

        keyword = Entry(self, bg= "White", fg='Black', background='White', width= 10, textvariable=self.keyword_text)
        keyword.place(x =450-50, y = 27)

        keyword_lbl = ttk.Label(self, text='Keyword: ')
        keyword_lbl.place(x = 330-50, y = 25)

        self.keyword_lbl_text = StringVar()

        keyword_cells = Entry(self, bg= "White", fg='Black', background='White', width= 10, textvariable=self.keyword_lbl_text)
        keyword_cells.place(x=450-50, y = 27+25)

        keyword_cells_lbl = ttk.Label(self, text='Cells: ')
        keyword_cells_lbl.place(x = 490, y = 25)

        self.reference_text = StringVar()

        reference = Entry(self, bg= "White", fg='Black', background='White', width= 10, textvariable=self.reference_text)
        reference.place(x=550, y = 27)

        reference_lbl = ttk.Label(self, text='Reference: ')
        reference_lbl.place(x = 330-50, y = 25+25)

        finish = Button(self, text="Next", font= ('Calabri', 17), borderwidth=1, relief="ridge",     
                        background= "#800020", foreground='White', activebackground="#a6022b" , activeforeground="White", cursor="hand2",
                        command= lambda: [controller.show_frame(Page2), ]
                        )
        finish.place(x = 1045, y = 645)

    def get_validation(self):
        return [self.keyword_text, self.keyword_lbl_text, self.reference_text]

class Page4(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.configure(bg=BACKGROUND_COLOR)

        #-----STYLE----#
        
        style = ttk.Style()
        style.configure('TLabelframe', background= SECONDARY_COLOR)
        style.configure('TLabelframe.Label', font =('Calibri', 15, 'bold'))
        style.configure('TLabelframe.Label', foreground = "White")
        style.configure('TLabelframe.Label', background = SECONDARY_COLOR )

        style.configure('TButton', font = ('Calibri', 15))
        style.configure('TButton', background = "#4733BF", )

        style.configure('TLabel', font = ('Calabri', 15))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)

        self.img = ImageTk.PhotoImage(Image.open("C:./Title Bar2.png"))
        # Use a label to place the image
        self.label = ttk.Label(self, image=self.img)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)

        self.loading = Label(self, text = "Loading .", font=('Calabri', 24), bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR)
        self.loading.place(x=500, y = 700/2)

    def get_loading_label(self):
        return self.loading
       
if __name__ == "__main__":
    app = tkinterApp()
    app.mainloop()

