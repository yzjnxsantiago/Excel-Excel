# Author: yzjnxsantiago
# Start Date: Tuesday, November 9, 2022

LARGEFONT =("Verdana", 35)
BACKGROUND_COLOR = "#262335"
SECONDARY_COLOR  = "#241b2f"
BUTTON_COLOR     = "#5a32fa"
BUTTON_HIGHLIGHT = "#7654ff"

#--------LIBRARIES--------#

from tkinter import *
import tkinter as tk
from tkinter import ttk
from zlib import Z_FIXED
from PIL import ImageTk, Image
from tkinter import filedialog
from building_blocks import *
import threading
import time

def browse_button(directory_label: Label, frame: Frame, sheet_frame: Frame):
    
    """_summary_ 

    Args:
        directory_label (Label): _description_
        frame (Frame): _description_
        sheet_frame (Frame): _description_
    """
    
    files_list = [] 
    source_sheets = []
    sheet_graphics = []
    sheet_variable = []

    x1 = 275
    x2 = 275
    y1 = 85+20
    y2 = y1 + 40

    placement_x = 290
    placement_y = 75

    filename = filedialog.askdirectory()
    directory_label.set(filename)

    files_list = find_files(".xlsx", filename)

    for i in range(len(files_list)):
        listBox(frame, x1,y1,x2,y2,files_list[i])
        y1 += 25
        y2 += 25

    source_sheets = id_sheets(filename)

    for i in range(len(source_sheets)):
            sheet_variable.append(StringVar())
            sheet_graphics.append(Checkbutton(sheet_frame, text=source_sheets[i]+ " ", variable=sheet_variable[i], font= ('Calabri', 10),
                                  background=BUTTON_HIGHLIGHT, foreground='white',  borderwidth = 1, relief="ridge",  height=2,
                                  activebackground=BUTTON_COLOR, activeforeground='White', selectcolor= BACKGROUND_COLOR ))
            sheet_graphics[i].place(x = placement_x, y = placement_y)
            sheet_graphics[i].update()
            sheet_graphics_width = sheet_graphics[i].winfo_width()
            placement_x = placement_x +  sheet_graphics_width + 15
            
            if placement_x > 1000:
                placement_x = 290
                placement_y += 50
        
class tkinterApp(tk.Tk):

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
         
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        self.geometry("1200x700")
        self.resizable("False", "False")
        self.title("Excel to Excel")
        
         
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
  
        # initializing frames to an empty array
        self.frames = {} 
  
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, Page1, Page2, Page3):
  
            frame = F(container, self)
  
            # initializing frame of that object from
            # startpage, page1, page2 respectively with
            # for loop
            self.frames[F] = frame
  
            frame.grid(row = 0, column = 0, sticky ="nsew")
  
        self.show_frame(StartPage)  
  
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def get_frame(self, cont):
        return self.frames[cont]

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

        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Explorer", height=698, width=250) 
        menu_bar.place(x = 0, y =0)

        #-Buttons-#

        create = Button(self, text ="Create", font=('Calibri', 15), fg= "white", bg="#5615DE", activebackground='#6017F9',
                            activeforeground='white',
                            command = lambda : controller.show_frame(Page1))
        create.place(x = 300, y = 20)



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

        style.configure('TLabel', font = ('Calabri', 15))
        style.configure('TLabel', foreground = 'White')
        style.configure('TLabel', background = SECONDARY_COLOR)

        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=698, width=250) 
        menu_bar.place(x = 0, y =0) 
        
        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection",borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                            activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y =30)
      
        #-----BROWSE FRAME-----#
        
        browse_frame = ttk.LabelFrame(self, text='Source', height = 100, width = 900)
        browse_frame.place(x = 275, y = 10)

        #-Buttons-#

        browse_source  = ttk.Button(self, text="Browse Files", 
                            command= lambda : browse_button(directorystr, self, controller.get_frame(Page2)))
        browse_source.place(x = 290, y = 35 )
        
        #-Labels-#

        directorystr = StringVar()
        source_directory = ttk.Label(self, textvariable=directorystr)
        source_directory.place(x = 420, y = 40)

        #------MAIN SECTION------#
        #-Buttons-#

        next = ttk.Button(self, text='Next', command = lambda : controller.show_frame(Page2))
        next.place(x = 1065, y = 650)

        #-Labels-#

        files = ttk.Label(self, text=' Source Files ', font=('Calabri', 12), borderwidth=2, relief="groove")
        files.place(x=275, y =115)


def listBox(frame, x1, y1, x2, y2, filename):

    if y2 > 605:
        return

    label = ttk.Label(frame, text=filename, font= ('Calabri', 10), borderwidth=2, relief="solid", width= 128)
    label.place(x = x2, y = y2)

  
# third window frame page2
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

        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=698, width=250) 
        menu_bar.place(x = 0, y =0)
        

        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection", borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y =30)

        #-------SHEET SELECTION-----#

        sheet_selection = ttk.LabelFrame(self, text='Sheet Selection', height = 250, width = 900)
        sheet_selection.place(x = 275, y = 20)

        #------MAIN SECTION-------#

        type_a = Button(self, text="Type a", font= ('Calabri', 15), borderwidth=1, relief="ridge", width= 15, height = 5, 
                        background= BUTTON_HIGHLIGHT, foreground='White', activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2",
                        )
        type_a.place(x=275, y = 550)

        type_b = Button(self, text="Type b", font= ('Calabri', 15), borderwidth=1, relief="ridge", width= 15, height = 5, 
                        background= BUTTON_HIGHLIGHT, foreground='White', activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2",
                        command= lambda: controller.show_frame(Page3) 
                        )
        type_b.place(x=275+200, y = 550)

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
        
        
        #-------MENU BAR-------#
        
        menu_bar = ttk.LabelFrame(self, text= "Menu", height=698, width=250) 
        menu_bar.place(x = 0, y =0)
        

        #-Buttons-#
        
        nav_source = Button(self, text="Source Directory Selection", borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", activebackground=BUTTON_COLOR , activeforeground="White",  
                            command = lambda: controller.show_frame(Page1))
        nav_source.place(x = 20 , y =30)

        #-------CELL SELECTION-------#

        sheet_selection = ttk.LabelFrame(self, text='Cell Selection', height = 90, width = 170)
        sheet_selection.place(x = 275, y = 20)

        source_cell = Entry(self, bg= BACKGROUND_COLOR, fg='White', background=BACKGROUND_COLOR, width= 5)
        source_cell.place(x = 290, y = 60)

        confirm_cell = Button(self,text="Confirm Cell", borderwidth=1, relief="ridge", background=BUTTON_HIGHLIGHT, foreground="White", 
                              activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2" 
                              )
        confirm_cell.place(x= 336, y = 58)

        #------CELL------#

        cell_frame = ttk.LabelFrame(self, text='Cell', height = 90, width = 90)
        cell_frame.place(x = 450, y = 20)

        cell = Button(self,text="A3", font =('Calabri', 15), borderwidth=1, relief="groove", background=BUTTON_HIGHLIGHT, foreground="White", 
                      activebackground=BUTTON_COLOR , activeforeground="White", cursor="hand2")
        cell.place(x=475, y=50)

        def drag_start(event):
            widget = event.widget
            widget.startX = event.x
            widget.startY = event.y

        def drag_motion(event):
            widget = event.widget
            x = widget.winfo_x() - cell.startX + event.x
            y = widget.winfo_y() - cell.startY + event.y
            widget.place(x=x,y=y)
            
        cell.bind('<Button-1>', drag_start)  
        cell.bind('<B1-Motion>', drag_motion)

   

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




class VerticalScrolledFrame(ttk.Frame):
    """A pure Tkinter scrollable frame that actually works!
    * Use the 'interior' attribute to place widgets inside the scrollable frame.
    * Construct and pack/place/grid normally.
    * This frame only allows vertical scrolling.
    """
    def __init__(self, parent, *args, **kw):
        ttk.Frame.__init__(self, parent, *args, **kw)

        # Create a canvas object and a vertical scrollbar for scrolling it.
        vscrollbar = ttk.Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=FALSE)
        canvas = tk.Canvas(self, bd=0, highlightthickness=0,
                           yscrollcommand=vscrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        vscrollbar.config(command=canvas.yview)

        # Reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

app = tkinterApp()
app.mainloop()

