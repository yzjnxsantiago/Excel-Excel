# Author: yzjnxsantiago
# Start Date: Monday, April 3, 2023

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
#from excel_to_excel import excel_excel
import threading

def browse_dir(directory_label: Label, frame: Frame, sheet_frame: Frame, controller):
    
    """_summary_ 
    This function is intended to let the user select a directory so that has
    simillarly formatted excel files

    """

    files_list = [] 
    source_sheets = []
    sheet_variable = controller.get_frame(WorkPage).get_sheet_variables()[0]
    text_sheet_variable = controller.get_frame(WorkPage).get_sheet_variables()[1]
    sheet_graphics = controller.get_frame(WorkPage).get_sheet_variables()[2]

    # Initial Placements

    x1 = 275
    x2 = 275
    y1 = 160 + 85+20
    y2 = y1 + 40

    placement_x = 290
    placement_y = 150 + 80

    dirname = filedialog.askdirectory()

    files_list = find_files(".xlsx", dirname)
    source_sheets = id_sheets(dirname)

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

    

def browse_file(directory_label):
    # Get user input for the destination file
    filename = filedialog.askopenfilename(title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                       ("All files",
                                                        "*.*")))
    last_slash_index = filename.rfind("/")
    short_filename = filename[last_slash_index+1:]
    directory_label.configure(text=short_filename)

class tkinterApp(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
         
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)

        width, height = self.winfo_screenwidth(), self.winfo_screenheight()

        self.geometry('%dx%d+0+0' % (width,height))
        self.resizable("True", "True")
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
        for F in (WorkPage, FinishedPage):
            
            
            frame = F(self.container, self)
  
            # initializing frame of that object from
            # startpage, page1, page2 respectively with
            # for loop
            self.frames[F] = frame
  
            frame.grid(row = 0, column = 0, sticky ="nsew")
            
  
        self.show_frame(WorkPage)  
  
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


class WorkPage(tk.Frame):

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

        self.img = Image.open("C:./Title Bar2.png")
        self.resizedimg = self.img.resize((1950,200), Image.ANTIALIAS)
        self.titlebar = ImageTk.PhotoImage(self.resizedimg)        
        self.label = ttk.Label(self, image=self.titlebar)
        self.label.image = self.img
        # Place the image at the top left of the screen
        self.label.place(x=0,y=0)
        
        #---LEFT FRAME---#
        # Menu Bar used for selecting the source and destination workbooks as well as well as
        # openning recently saved files
        self.left_frame = Frame(self, width=100, height=900, bg= SECONDARY_COLOR)
        self.left_frame.place(x=0, y=202)
        
        #---CENTER FRAME---#
        # create the canvas and add it to the window
        canvas = Canvas(self, width=1330, height=250, bg=SECONDARY_COLOR, highlightthickness=0)
        canvas.place(x=530, y=220)

        # create a scrollable frame inside the canvas
        scrollable_frame = Frame(canvas, bg=SECONDARY_COLOR)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # add a horizontal scrollbar to the canvas
        scrollbar = Scrollbar(self, orient="horizontal", command=canvas.xview)
        scrollbar.place(x=530, y=520-50, width=1330)
        canvas.configure(xscrollcommand=scrollbar.set)

        # configure the canvas to resize with the window
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", on_configure)

        # set the size of the scrollable frame
        scrollable_frame.update_idletasks()
        scrollable_frame.config(width=scrollable_frame.winfo_reqwidth(), height=scrollable_frame.winfo_reqheight())

        # add some extra widgets to the bottom of the canvas
        spacer = Frame(canvas, height=20, bg=SECONDARY_COLOR)
        canvas.create_window((0, 0), window=spacer, anchor="nw")

        # Source Workbook Button #

        src_collect_btn = Button(self.left_frame, text ="Xs", font=('Calibri', 25), fg= "white", bg="#5615DE", activebackground='#6017F9',
                            activeforeground='white', width=4, command = lambda :browse_dir(src_lbl))
        src_collect_btn.place(x=10,y=15)
        
        # Destination Workbook Button #

        des_collect_btn = Button(self.left_frame, text ="X", font=('Calibri', 25), fg= "white", bg="#5615DE", activebackground='#6017F9',
                            activeforeground='white', width=4, command= lambda: browse_file(des_lbl))
        des_collect_btn.place(x=10,y=15+90)

        # Open File Button #

        open_file_btn = Button(self.left_frame, text ="+", font=('Calibri', 25), fg= "white", bg="#5615DE", activebackground='#6017F9',
                            activeforeground='white', width=4)
        open_file_btn.place(x=10,y=15+90*2)
    
        # Source and Destination Labels #

        src_scribe_lbl = Label(self, text = "Source: ", font=('Calabri', 24), bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR)
        src_scribe_lbl.place(x = 110, y = 230)

        src_lbl = Label(self, font=('Calabri', 11), bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR)
        src_lbl.place(x = 110+165, y = 250)

        des_scribe_lbl = Label(self, text = "Destination: ", font=('Calabri', 22), bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR)
        des_scribe_lbl.place(x = 110, y = 230+90)

        des_lbl = Label(self, font=('Calabri', 11), bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR)
        des_lbl.place(x = 110+165, y = 250+82)

        #----------VARIABLES-----------#
        
        self.sheet_variable = []
        self.text_sheet_variable = []
        self.sheet_checkboxes = []


    def get_sheet_variables(self):
        return [self.sheet_variable, self.text_sheet_variable, self.sheet_checkboxes]

class FinishedPage(tk.Frame):
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