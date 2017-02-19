import tkinter as tk
from tkinter.filedialog import *
import time, os, collections, datetime, json
from sheet import Spreadsheet
from top_wins import *

#from trial import Multi_box

ALL = N+S+W+E

class Application(Frame):
    
              
    def __init__(self, master=None):
        '''Create a 'master' frame of 1 row x 1 column'''
        tk.Frame.__init__(self, master,
                       bg = 'snow3',
                       bd = 4,
                       highlightbackground = 'grey')
        self.HEADERS = ['First',
                        'Middle',
                        'Last',
                        'Street',
                        'Apt#',
                        'City',
                        'Zip',
                        'Email',
                        'Occupation',
                        'Employer',
                        'Filed By',
                        'Rolodex',
                        'Member Signed',
                        'Cope Signed',
                        'Date Signed',
                        'Contribution',]
	
        self.DEFAULT_EXTENSIONS = [("Spreadsheet files", "*.csv\
                                                          *.xls\
                                                          *.xlsb\
                                                          *.ods"),
                                    ("Template files", "*.tplate"),
                                    ("HTML files", "*.html;*.htm"),
                                    ("All files", "*.*") ]

        #Create a CSV Spreadsheet object    
        self.sheet = Spreadsheet(self.HEADERS)
    
        #Default Path used for tkinter open/save_as widgets
        self.DEFAULT_PATH = os.getcwd()

        #GUI Set Up
        self.master.rowconfigure(0, weight = 1)
        self.master.columnconfigure(0, weight = 1)
        self.grid(sticky = W+E+S+N)

        #Key Bindings
        self.master.bind('<Return> ', self.gather)
        #Drop down menu shortcuts
        self.master.bind('<Control_L><n>', self.new)
        self.master.bind('<Control_R><n>', self.new)
        self.master.bind('<Control_L><o>', self.load)
        self.master.bind('<Control_R><o>', self.load)
        self.master.bind('<Control_L><m>', self.merge)
        self.master.bind('<Control_R><m>', self.merge)
        self.master.bind('<Control_L><f>', self.find)
        self.master.bind('<Control_R><f>', self.find)
        self.master.bind('<Control_L><s>', self.file_save)
        self.master.bind('<Control_R><s>', self.file_save)
        self.master.bind('<Control_L>,Shift_L><S>', self.save_as)
        self.master.bind('<Control_L><Shift_R><S>', self.save_as)
        self.master.bind('<Control_R><Shift_L><S>', self.save_as)
        self.master.bind('<Control_R><Shift_R><S>', self.save_as)
        self.master.bind('<Control_L>,Shift_L><s>', self.save_as)
        self.master.bind('<Control_L><Shift_R><s>', self.save_as)
        self.master.bind('<Control_R><Shift_L><s>', self.save_as)
        self.master.bind('<Control_R><Shift_R><s>', self.save_as)
        self.master.bind('<Control_L><i>', self.read_out)
        self.master.bind('<Control_R><i>', self.read_out)
        self.master.bind('<Control_L><r>', self.remove_row)
        self.master.bind('<Control_R><r>', self.remove_row)
        self.master.bind('<Control_L><c>', self.remove_column)
        self.master.bind('<Control_R><c>', self.remove_column)
        self.master.bind('<Control_L><e>', self.edit_cell)
        self.master.bind('<Control_R><e>', self.edit_cell)
        self.master.bind('<Control_L><h>', self.help)
        self.master.bind('<Control_R><h>', self.help)
        self.master.bind('<Control_L><a>', self.about)
        self.master.bind('<Control_R><a>', self.about)

        #The Widgets
        self.createWidgets()


        #Establish and set highest know Rolodex
        self.ROLODEX_SERIAL = 'Serialed_rolodex'
        dex = max(self.load_rolodex(),self.sheet.ROLODEX)
        self.sheet.ROLODEX = dex
        self.rolodex_set(dex)


        self.ENTRIES = [self.First_name_entry,
                        self.Middle_name_entry,
                        self.Last_name_entry,
                        self.Street_entry,
                        self.Apt_entry,
                        self.City_entry,
                        self.Zip_entry,
                        self.Email_entry,
                        self.Occupation_entry,
                        self.Employer_entry,
                        self.Filed_by_entry,
                        self.Rolodex_entry,
                        self.Member_but_val,
                        self.Cope_but_val,
                        self.Signed_entry,
                        self.Contribute_val
                        ]
        
    def createWidgets(self):
        '''Create and add widgets to 'master' frame'''
        for r in range(30):
            self.rowconfigure(r, weight = 1)
        for  c in range(60):
            self.columnconfigure(c, weight =1)

        #Menu Bar
        menu = Menu(root)
        root.config(menu=menu)
        #File Menu
        filemenu = Menu(menu)
        menu.add_cascade(label = 'File', menu=filemenu)
        #New
        filemenu.add_command(label='New               Ctrl+n',
                             command= self.new)
        #Load
        filemenu.add_command(label = 'Load               Ctrl+o',
                             command = self.load)
        #Merge
        filemenu.add_command(label = 'Merge            Ctrl+m',
                            command = self.merge)

        #Find
        filemenu.add_command(label = 'Find                Ctrl+f',
                             command = self.find)

        #Save
        filemenu.add_command(label = 'Save                Ctrl+s',
                             command = self.file_save)
        #Save as
        filemenu.add_command(label = 'Save As...        Ctrl+Shift+s',
                             command = self.save_as)
        #View Menu
        viewmenu = Menu(menu)
        menu.add_cascade(label = 'View', menu=viewmenu)

        #Readout
        viewmenu.add_command(label = 'Read Out        Ctrl+i',
                             command = self.read_out)
        
        #Edit Menu
        editmenu = Menu(menu)
        menu.add_cascade(label = 'Edit', menu=editmenu)
        
        #Remove Row
        editmenu.add_command(label = 'Remove Row               Ctrl+r',
                             command = self.remove_row)

        #Remove Column
        editmenu.add_command(label = 'Remove Column         Ctrl+c',
                             command = self.remove_column)

        #Edit Cell
        editmenu.add_command(label = 'Edit Cell                        Ctrl+e',
                             command = self.edit_cell)
        

        #Help Menu
        helpmenu = Menu(menu)
        menu.add_cascade(label = "Help", menu=helpmenu)

        #Help
        helpmenu.add_command(label = 'Help...        Ctrl+h',
                             command = self.help)
        #About
        helpmenu.add_command(label = 'About...',
                             command = self.about)
        
        #CONSTANT COLOR VARIABLES
        BG = 'snow3'
        FG = 'black'
        LINE = 'white'
        FRAME = 'grey'
        BG_BUTTON = BG
        READOUT = 'grey'

        #Padding Values
        PADX = 10
        PADY = 10
        
        #LEFT Frame
        FLEFT = Frame(self, bg = BG)
        FLEFT.grid(row = 0,
                   column = 0,
                   rowspan = 31,
                   columnspan = 29,
                   sticky = ALL)
        #RIGHT Frame
        FRIGHT = Frame(self, bg = BG)
        FRIGHT.grid(row = 0,
                    column = 30,
                    rowspan = 31,
                    columnspan = 29,
                    sticky = ALL)

        #Name
        #First Name
        self.First_name = Frame(FLEFT, bg = BG)
        self.First_name_entry = Entry(self.First_name)
        self.First_name_entry.pack()
        self.First_name_entry.focus()
        self.First_name_label = Label(self.First_name,
                                      text = "First Name",
                                      bg = BG,
                                      fg = FG)
        
        self.First_name_label.pack(side = LEFT)
        self.First_name.grid(row = 1,
                             column = 0,
                             columnspan = 2,
                             padx = PADX,
                             pady = PADY,
                             sticky = W)

        #Middle Name
        self.Middle_name = Frame(FLEFT, bg = BG)
        self.Middle_name_entry = Entry(self.Middle_name)
        self.Middle_name_entry.pack()
        self.Middle_name_label = Label(self.Middle_name,
                                      text = "Middle Name",
                                      bg = BG,
                                      fg = FG)
        
        self.Middle_name_label.pack(side = LEFT)
        self.Middle_name.grid(row = 1,
                              column = 2,
                              columnspan = 2,
                              padx = PADX,
                              pady = PADY,
                              sticky = W)

        #Last Name
        self.Last_name = Frame(FLEFT, bg = BG)
        self.Last_name_entry = Entry(self.Last_name)
        self.Last_name_entry.pack()
        self.Last_name_label = Label(self.Last_name,
                                     text = "Last Name",
                                     bg = BG,
                                     fg = FG)
        
        self.Last_name_label.pack(side = LEFT)
        self.Last_name.grid(row = 1,
                            column = 5,
                            columnspan = 2,
                            padx = PADX,
                            pady = PADY,
                            sticky = W)

        #Address
        #Street Address
        self.Street = Frame(FLEFT, bg = BG)
        self.Street_entry = Entry(self.Street)
        self.Street_entry.pack()
        self.Street_label = Label(self.Street,
                                      text = "Street Address",
                                      bg = BG,
                                      fg = FG)
        
        self.Street_label.pack(side = LEFT)
        self.Street.grid(row = 2,
                             column = 0,
                             columnspan = 2,
                             padx = PADX,
                             pady = PADY,
                             sticky = W)

        #Apt
        self.Apt = Frame(FLEFT, bg = BG)
        self.Apt_entry = Entry(self.Apt,
                               width = 5)
        self.Apt_entry.pack()
        self.Apt_label = Label(self.Apt,
                                text = "Apt #",
                                bg = BG,
                                fg = FG)
        
        self.Apt_label.pack(side = LEFT)
        self.Apt.grid(row = 2,
                              column = 2,
                              columnspan = 2,
                              padx = PADX,
                              pady = PADY,
                              sticky = W)

        #City
        self.City = Frame(FLEFT, bg = BG)
        self.City_entry = Entry(self.City,
                                width = 30)
        self.City_entry.pack()
        self.City_label = Label(self.City,
                                text = "City",
                                bg = BG,
                                fg = FG)
        
        self.City_label.pack(side = LEFT)
        self.City.grid(row = 3,
                       column = 0,
                       columnspan = 3,
                       padx = PADX,
                       pady = PADY,
                       sticky = W)


        #Zipcode
        self.Zip = Frame(FLEFT, bg = BG)
        self.Zip_entry = Entry(self.Zip,
                                width = 8)
        self.Zip_entry.pack()
        self.Zip_label = Label(self.Zip,
                                text = "Zip",
                                bg = BG,
                                fg = FG)
        
        self.Zip_label.pack(side = LEFT)
        self.Zip.grid(row = 3,
                      column = 3,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Email
        self.Email = Frame(FLEFT, bg = BG)
        self.Email_entry = Entry(self.Email,
                                width = 30)
        self.Email_entry.pack()
        self.Email_label = Label(self.Email,
                                text = "Email",
                                bg = BG,
                                fg = FG)
        
        self.Email_label.pack(side = LEFT)
        self.Email.grid(row = 5,
                      column = 0,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Occupation
        self.Occupation = Frame(FLEFT, bg = BG)
        self.Occupation_entry = Entry(self.Occupation,
                                width = 30)
        self.Occupation_entry.pack()
        self.Occupation_label = Label(self.Occupation,
                                text = "Occupation",
                                bg = BG,
                                fg = FG)
        
        self.Occupation_label.pack(side = LEFT)
        self.Occupation.grid(row = 10,
                      column = 0,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Employer
        self.Employer = Frame(FLEFT, bg = BG)
        self.Employer_entry = Entry(self.Employer,
                                width = 30)
        self.Employer_entry.pack()
        self.Employer_label = Label(self.Employer,
                                text = "Employer",
                                bg = BG,
                                fg = FG)
        
        self.Employer_label.pack(side = LEFT)
        self.Employer.grid(row = 10,
                      column = 3,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Line Break
        self.line_break = Frame(FLEFT, bg = BG)
        self.line_break_label = Label(self.line_break,
                                text = "_"*80,
                                bg = BG,
                                fg = LINE)
        self.line_break_label.pack(side = LEFT)
        self.line_break.grid(row = 12,
                      column = 0,
                      columnspan = 9,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Filed By
        self.Filed_by = Frame(FLEFT, bg = BG)
        self.Filed_by_entry = Entry(self.Filed_by,
                                width = 30)
        self.Filed_by_entry.pack()
        self.Filed_by_label = Label(self.Filed_by,
                                text = "Filed By",
                                bg = BG,
                                fg = FG)
        
        self.Filed_by_label.pack(side = LEFT)
        self.Filed_by.grid(row = 13,
                      column = 0,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)

        #Rolodex 
        self.PICKLED_ROLODEX = Frame(FLEFT, bg = BG)
        self.Rolodex_entry = Entry(self.PICKLED_ROLODEX,
                                width = 30)
        self.Rolodex_entry.pack()
        
        self.Rolodex_entry.insert(0, self.sheet.ROLODEX)
            
        self.Rolodex_entry.config(state = 'disabled')
        self.Rolodex_label = Label(self.PICKLED_ROLODEX,
                                text = "Rolodex",
                                bg = BG,
                                fg = FG)
        
        self.Rolodex_label.pack(side = LEFT)
        self.PICKLED_ROLODEX.grid(row = 13,
                      column = 3,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)
        
        #Member Signed Radio Buttons
        self.Butts = Frame(FRIGHT,
                           bg = BG)
        self.Butts.grid(row = 0,
                                column = 1,
                                rowspan = 3,
                                columnspan = 16,
                                sticky = ALL)
        
        self.Member_signed = Frame(self.Butts,
                                   bg = BG,
                                   highlightthickness = 1,
                                   highlightbackground = FRAME)
        self.Member_signed.grid(row = 0,
                                column = 0,
                                columnspan = 8,
                                sticky = E)
        
        self.Member_signed_label = Label(self.Member_signed,
                                         text = "Member Card Signed?",
                                         bg = BG,
                                         fg = FG).pack(side = TOP)        
        self.Member_but = Frame(self.Member_signed, bg = BG)
        self.Member_but.pack(side = RIGHT)
        self.Member_but_val = IntVar()
        self.Member_but_val.set(0)
        self.Member_yes = Radiobutton(self.Member_signed,
                                    text = "Yes",
                                    variable = self.Member_but_val,
                                    value = 1,
                                    bg = BG,
                                    fg = FG)
        self.Member_yes.pack(side = LEFT)
        self.Member_no = Radiobutton(self.Member_signed,
                                    text = "No",
                                    variable = self.Member_but_val,
                                    value = 0,
                                    bg = BG,
                                    fg = FG)
        self.Member_no.pack(side = LEFT)

        #C.O.P.E Card Signed Radio Buttons
        self.Cope_signed = Frame(self.Butts,
                                 bg = BG,
                                 highlightthickness = 1,
                                 highlightbackground = FRAME)
        self.Cope_signed.grid(row = 0,
                                column = 11,
                                rowspan = 3,
                                columnspan = 8,
                                sticky = E)
        self.Cope_signed_label = Label(self.Cope_signed,
                                         text = "C.O.P.E Card Signed",
                                         bg = BG,
                                         fg = FG).pack(side = TOP)        
        self.Cope_but = Frame(self.Cope_signed, bg = BG)
        self.Cope_but.pack(side = RIGHT)
        self.Cope_but_val = IntVar()
        self.Cope_but_val.set(0)
        self.Cope_yes = Radiobutton(self.Cope_signed,
                                    text = "Yes",
                                    variable = self.Cope_but_val,
                                    value = 1,
                                    bg = BG,
                                    fg = FG)
        self.Cope_yes.pack(side = LEFT)
        self.Cope_no = Radiobutton(self.Cope_signed,
                                    text = "No",
                                    variable = self.Cope_but_val,
                                    value = 0,
                                    bg = BG,
                                    fg = FG)
        self.Cope_no.pack(side = LEFT)


        #Date Member Signed
        self.Date_entry = Frame(FRIGHT, bg = BG)
        self.Signed_entry = Entry(self.Date_entry,
                                width = 18)
        self.Signed_entry.pack(side = LEFT)
        self.Date_entry.grid(row = 10,
                      column = 1,
                      columnspan = 3,
                      padx = PADX,
                      pady = PADY,
                      sticky = W)
        self.Signed = Frame(FRIGHT, bg=BG)
        self.Signed_label = Label(self.Signed,
                                  text = "Date Signed   (mm/dd/yyyy)\n",
                                  bg = BG,
                                  fg = FG)
        
        self.Signed_label.pack(side = RIGHT)
        self.Signed.grid(row = 11,
                         column = 1,
                         columnspan = 3,
                         padx = PADX,
                         sticky = W)

        #contribution amount in Radio Buttons
        self.Contribute = Frame(FRIGHT, bg= BG)
        self.Contribute.grid(row = 12,
                             column = 0,
                             columnspan =10,
                             sticky = W)
        self.Contribute_val = IntVar()
        self.Contribute_val.set(0)
        self.Contrib_label = Frame(FRIGHT, bg = BG)
        self.Contribute_label = Label(self.Contrib_label,
                                      text = '\nMonthly Contribution Amount',
                                      bg = BG,
                                      fg = FG).pack(side=LEFT)
        self.Contrib_label.grid(row = 13,
                                column = 1)
         

        self.Contribute_none = Radiobutton(self.Contribute,
                                          text = '-  0',
                                          variable = self.Contribute_val,
                                          value = 0,
                                          bg = BG,
                                          fg = FG).pack(side=LEFT)
        self.Contribute_seven = Radiobutton(self.Contribute,
                                          text = '-  7',
                                          variable = self.Contribute_val,
                                          value = 7,
                                          bg = BG,
                                          fg = FG).pack(side=LEFT)
        
        self.Contribute_ten = Radiobutton(self.Contribute,
                                          text = '-10',
                                          variable = self.Contribute_val,
                                          value = 10,
                                          bg = BG,
                                          fg = FG).pack(side=LEFT)
        self.Contribute_fifteen = Radiobutton(self.Contribute,
                                          text = '-15',
                                          variable = self.Contribute_val,
                                          value = 15,
                                          bg = BG,
                                              fg = FG).pack(side=LEFT)
        self.Contribute_twenty = Radiobutton(self.Contribute,
                                          text = '-20',
                                          variable = self.Contribute_val,
                                          value = 20,
                                          bg = BG,
                                          fg =FG).pack(side=LEFT)

        ##Cheater line##
        #invisible line to push button down
        '''
        Was attempting to get the 'Submit' button to sit lower in the corner
        than it currently was, it wasn't cooperating so I hacked this in to
        force it down.
        '''
        self.Spacer_Frame = Frame(FRIGHT, bg = BG)
        self.Spacer_Frame.grid(row = 29,
                             column = 4)
        self.Spacer = Label(self.Spacer_Frame,
                                 text = '\n\n',
                                 bg = BG)
        self.Spacer.pack()

        #Submit Button
        '''
        Store entry field values in a dict of dicts, rolodex is inner dict key:
        cards = {rolodex : {entry_fields:values}}
        '''
        self.Submit = Frame(FRIGHT, bg = BG)
        self.Submit.grid(row=30,
                         column=3,
                         columnspan = 10,
                         sticky= ALL)
        self.submit_but = Button(self.Submit,
                                 text="Next Card >>",
                                 command = self.gather,
                                 height = 2,
                                 bg = BG_BUTTON)
                                 
        self.submit_but.pack(side = RIGHT)

        #Readout/Display #1
        self.Display_1 = Frame(FRIGHT, bg = BG)
        self.Display_1_label = Label(self.Display_1,
                                     bg = BG,
                                     fg = READOUT)
        self.Display_1_label.pack(side = LEFT)
        self.Display_1.grid(row = 20,
                      column = 0,
                      columnspan = 9,
                      rowspan = 4,
                      pady = PADY*2,
                      sticky = W)
        
        #Readout/Display #2
        self.Display_2 = Frame(FRIGHT, bg = BG)
        self.Display_2_label = Label(self.Display_2,
                                     bg = BG,
                                     fg = READOUT)
        self.Display_2_label.pack(side = LEFT)
        self.Display_2.grid(row = 24,
                      column = 0,
                      columnspan = 9,
                      rowspan  = 2,
                      pady = PADY*2,
                      sticky = W)

    #Dropdown menu callbacks
    def new(self,event = None):
        '''
        Create a new blank Spreadsheet object w/ headers, if file name not
        changed with save_as, (n) will be appended to file name.
        '''
        
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')
        self.sheet = Spreadsheet(self.HEADERS)
        self.clear_display()
        rolodex = self.load_rolodex()
        self.sheet.ROLODEX = rolodex
        self.rolodex_set(rolodex)
        


    def load(self, event = None):
        '''
        Reset Spreadsheet objects read/WRITE_DICT with data rows from an
        existing file.
        '''
        old_dex = self.load_rolodex()
        path, fn = os.path.split(self.browse())
        full_path = os.path.join(path, fn)
        self.sheet.load(full_path)
        new_dex = self.sheet.ROLODEX
        if new_dex > old_dex:
            self.sheet.ROLODEX += 1  
            self.rolodex_set(self.sheet.ROLODEX)
        self.que_display()
        self.fn_display(msg = 'Loaded')
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')
            
    def merge(self,event = None):
        '''
        Create a new merged sheet from current working sheet and a second
        existing sheet; new sheet auto-named as [fn]_merged.csv and saves in
        working directory, if duplicate rolodex found, prompted to change
        or ovrwrite entry with overlaping rolodex number
        '''
        #Assure uniform label scheme
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')

        fn = self.browse()
    
        msg = self.sheet.merge(fn)
        
        dex = self.load_rolodex()
        self.rolodex_set(dex)
        self.sheet.ROLODEX = dex
        self.store_rolodex()
        self.msg_display(2, msg)
        self.que_display()

    def find(self,event = None): 
        '''
        Open a child window with two entry fields, one for a header value,
        one for a cell value, return all rows mathcing provided values;
        #IMPORTANT: Overwrites current sheet
        '''
        found = Double_entry('Find','Header', 'Value', 'Search')
        if found.head and found.val:
            msg = self.sheet.find(found.head.strip(), found.val.strip())
            self.clear_display()
            
            
            self.que_display()
            self.msg_display(2, msg)
            self.First_name_label.config(fg = 'black')
            self.Last_name_label.config(fg = 'black')

    def read_out(self, event = None):
        print("READ OUT")
        sheet = self.sheet.WRITE_DICT
        fn = self.sheet.FN
        od = collections.OrderedDict(sorted(sheet.items(),
                                            key = lambda t: t[0]))

        Read_out(self.HEADERS, od, fn)
   

    def file_save(self,event = None, mode = 'w'):
        '''
        Create and Save spreadsheet in working directory with working file name,
        file name default is current date
        '''
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')

        self.sheet.write(self.sheet.WRITE_DICT, mode)
        self.store_rolodex()
        #Display File Name in blank label space
        self.fn_display(msg = 'Saved')

    def save_as(self,event = None, mode = 'w'):
        '''
        Collects entry field data, open a Save_As file dialog box,
        Overwrites existing data if file alread exists
        '''
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')

        fn = asksaveasfile(parent = self.master,
                           defaultextension = '.csv',
                           initialdir = self.DEFAULT_PATH,
                           filetypes = self.DEFAULT_EXTENSIONS)
        if fn:
            fn = os.path.abspath(fn.name)
            self.sheet.FN = fn
            self.sheet.write(self.sheet.WRITE_DICT, mode = mode)
            self.store_rolodex()
            self.fn_display(msg = 'Saved')
        

    def remove_row(self, event = None):
        '''
        Remove single, or multiple rows based on a combination of Rolodex Id's,
        and Header = Value pairs.
        '''
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')
        remove = Triple_entry('Remove Row',
                              'Rolodex',
                              'Header',
                              'Cell Value',
                              'Remove Row')
        try:
            rolodex, header, value = remove.dex, remove.head, remove.val
            if rolodex == '' and header == '' and value == '':   
                msg = "***Missing Required Fields***"
                self.msg_display(1, msg, color = 'red')
            else:
                keys = self.sheet.remove_row(rolodex, header, value)
                            
            if len(keys) < 1:
                msg = '***No Matching Entries Found***'
                self.msg_display(1, msg, color = 'red')
            else:
                self.clear_display()
                msg = 'Removed  {}  from {}'.format(keys, self.sheet.FN)
                self.que_display()
                self.msg_display(2, msg)

        except AttributeError:
            #AttributeError if window X'd out without values
            pass

    def remove_column(self,event = None):
        '''
        Remove all data in the column under a given header
        '''
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')
        header = Single_entry('Remove Column','Header', 'Remove Column')
        
        try:
            self.sheet.remove_col(header.head.title())
        except AttributeError:
            #AttributeError if window X'd out without values
            pass
            
            
    def edit_cell(self, event = None):
        '''
        Alter the value of a single cell under a given header,
        accepts multiple Rolodex Values seperated by a comma
        '''
        self.First_name_label.config(fg = 'black')
        self.Last_name_label.config(fg = 'black')
        edit = Triple_entry('Edit Cell',
                            'Rolodex',
                            'Header',
                            'New Cell Value',
                            'Edit')
        try:
            self.sheet.edit_cell(edit.dex, edit.head.title(), edit.val)
            color = 'grey'
            msg = 'Altered Rows: {}\nHeader: {}\nNew Value: {}'.format(
                                                                 edit.dex,
                                                                 edit.head.title(),
                                                                 edit.val)
            if edit.head.title() not in self.HEADERS:
                color = 'red'
                msg = '***Invalid Header given***'
            self.msg_display(1, msg, color)

    
        except AttributeError:
            #AttribueError if window X'd with no values
            pass

            
        
    def help(self, event = None):
        categories = ['Merge',
                      'Find',
                      'Remove Row',
                      'Remove Column',
                      'Edit Cell',
                      'Read Out']
        func = Help_window(sorted(categories), 'Search')
        

    def about(self, event = None):
        msg = '''
        Card Catalogue 2.0

        Created and Designed By Kelly Wyss
        kelly.wyss@yahoo.com

        Read Out window courtesy of PythonMegaWidgets:\t
        http://pmw.sourceforge.net
        ''' 
        '''
        Deisgned and Intended for use by the office staff of a specific
        labor union office with a high rotation of temps to help streamline
        the process of data entry of union membership cards.
        '''
        
        Basic_display('About Card Catalogue', msg)

    #Private Functions
    def que_display(self):
        '''
        Display numer of cards qued, waiting to be written to Spreadsheet file,
        '''
        msg = "Total Cards Qued:{0:>7}".format(
                                        len(self.sheet.WRITE_DICT.keys()))
        self.msg_display(1, msg, color = 'black')
        
    def fn_display(self, msg = None):
        '''
        Display current Spreadsheet file name in lower display space
        '''
        text = msg + ": {0:60}".format(self.sheet.FN)
        self.msg_display(2, text, color = 'black')
        
    def msg_display(self, display, msg, color = 'grey'):
        '''
        display a message in one of two defined display spaces
        '''
        
        if display == 1:
            self.Display_1_label.config(text = msg, fg = color)
            
        if display == 2:
            self.Display_2_label.config(text = msg, fg = color)
            

    def clear_display(self):
        '''
        Clear both  display spaces
        '''
        self.Display_1_label.config(text = "")
        self.Display_2_label.config(text = "")
                
    def browse(self):
        '''
        Open a file browser widget, return a path/file_name
        '''
        fn = askopenfilename(filetypes = self.DEFAULT_EXTENSIONS,
                             initialdir = self.DEFAULT_PATH)
        if fn:
            fname = os.path.abspath(fn)

        return fn

    def rolodex_set(self, value):
        '''
        Enable rolodex entry field, set its value, and disable field again
        '''
        self.Rolodex_entry.config(state = 'normal')
        self.Rolodex_entry.delete(0, END)
        self.Rolodex_entry.insert(0, value)
        self.Rolodex_entry.config(state = 'disabled')

    def store_rolodex(self):
        '''
        Compare Spreadsheet's Roldex against Rolodex already serialized if any,
        serialize larger of the two with JSON
        '''

        rolodex = self.load_rolodex()
        with open(self.ROLODEX_SERIAL, 'w') as f:
            json.dump(rolodex, f)

    def load_rolodex(self):
        '''
        Compare Spreadsheet's Roldex against serialized Rolodex if any,
        return larger of the two
        '''
        fn= self.ROLODEX_SERIAL
        if os.path.isfile(fn):
            with open(fn,'r') as f:
                dex = json.load(f)
                rolodex = max(dex, self.sheet.ROLODEX)
        else:
            rolodex = self.sheet.ROLODEX

        return rolodex

              
    def gather(self, event = None):
        '''
        Create and return a dict of header:values for all entries, buttons
        and boxes
        '''
        entries = []
        #Unlock Rolodex widget, to read and increase its value 
        self.Rolodex_entry.config(state = 'normal')

        #Check Date signed not in future
        DATE_FORMAT = '%m/%d/%Y'
        date = time.strftime(DATE_FORMAT)
        signed = self.Signed_entry.get()

    
        #retrieve and clear all entries
        for entry in self.ENTRIES:
            entries.append(entry.get())
            try:
                entry.delete(0, END)
            except AttributeError:
                entry.set(0)
        #Convert Binary Radio button values to yes/no
        if entries[12] == 0:
            entries[12] = 'no'
        elif entries[12] == 1:
            entries[12] = 'yes'
        if entries[13] == 0:
            entries[13] = 'no'
        elif entries[13] == 1:
            entries[13] ='yes'
        #convert contirbution value to <str>
        entries[-1] = str(entries[-1])
  
        #Increase Rolodex by 1 if Card has a minimum of a first and last name
        first, last = entries[:3:2]
        if first == '' or last == '':
            self.First_name_entry.focus()
            self.rolodex_set(self.load_rolodex())
            self.First_name_label.config(fg = 'red')
            self.Last_name_label.config(fg = 'red')
            msg = '***Missing Required Fields***'
            self.clear_display()
            self.msg_display(2, msg, color = 'red')
        elif signed > self.sheet.DATE:
            self.Signed_entry.focus()
            self.clear_display()
            msg = '***Invalid Date Given: {}***'.format(signed)
            self.msg_display(1, msg, color = 'red')
            self.Signed_label.config(fg = 'red')
            
        else:
            signed <= self.sheet.DATE
            self.Signed_label.config(fg ='black')
            self.que_display()
            self.First_name_label.config(fg = 'black')
            self.Last_name_label.config(fg = 'black')
            self.msg_display(2, msg = '')
            dex = self.load_rolodex()
            self.sheet.ROLODEX += 1
            #auto-fill the rolodex number with the new value
            #and re-disable widget
            self.rolodex_set(self.sheet.ROLODEX)

            #Dict comp to turn entries into {header:value}
            card= {key:value for key, value in zip(self.HEADERS, entries)}        

            #Add card to Dict Que to be written to spreadsheet,
            #Rolodex is unique ID
            self.sheet.WRITE_DICT[card['Rolodex']] = card
            #Display Total cards Qued in an initially blank label space
            self.que_display()
            self.First_name_entry.focus()

    

   
        


root = Tk()
root.title("Card Catalogue 2.0")
app = Application(master=root)
app.mainloop()
