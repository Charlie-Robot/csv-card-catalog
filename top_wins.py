#from tkinter import *
import string
import tkinter as tk
import Pmw
import collections
import help_text


class Top:
    '''
    Custom parent class for top_level child windows to pull funcs from
    '''

    def but_funcs(self, *funcs):
        '''
        Call multiple functions with a single button click,
        1)get rolodex, header and value entry values
        2)destroy toplevel window
        '''
        def commands(*args, **kwargs):
            for f in funcs:
                f(*args, **kwargs)
        return commands

    def get_all(self, *args):
        '''
        get() toplevel entry field values
        '''
        try:
            self.dex = self.rolodex.get()
            self.head = self.header.get()
            self.val = self.value.get()
            return (self.dex, self.head, self.val)
        except AttributeError:
            pass

    def get_both(self, *args):
        try:
            self.head = self.header.get()
            self.val = self.value.get()
            return (self.head, self.val)
        except AttributeError:
            pass

    def get_single(self, *args):
        try:
            self.head = self.header.get()
            return self.head
        except AttributeError:
            pass

    def destroy(self, *args):
        '''
        Destroy toplevel window
        '''
        self.win.destroy()


    
class Single_entry(Top):
    '''
    Toplevel window wtih one button, returns a single entryfield value
    '''
    def __init__(self,
                 title,
                 label,
                 button_label):
        self.win = tk.Toplevel()
        self.win.title(title)
        self.label = label
        self.button_label = button_label
        self.master = tk.Frame(self.win)
        self.master.pack()
        self.upper = tk.Frame(self.master)
        self.upper.pack()
        self.lower = tk.Frame(self.master)
        self.lower.pack()
        self.border = 'grey'
        self.create_widgets()
        
    def create_widgets(self):
        self.entry_frame = tk.Frame(self.upper)
        
        self.header = tk.Entry(self.entry_frame,
                             highlightthickness = 1,
                             highlightbackground = self.border)
        self.header.grid(row=6, column = 0)
        self.header.focus()
        self.head_lab = tk.Label(self.entry_frame,text = self.label)
        self.head_lab.grid(row=7, column = 0)
        self.entry_frame.pack()
        self.button_frame = tk.Frame(self.lower)
        width = max(10, len(self.button_label))
        self.but = tk.Button(self.button_frame,
                          text = self.button_label,
                          width = width,
                          command = self.but_funcs(self.get_single, self.destroy))
        self.but.grid(row=8,column = 1)
        self.but.bind("<Return>", self.but_funcs(self.get_single, self.destroy))   
        self.button_frame.pack()
        
        self.win.wait_window()



class Double_entry(Top):
    '''
    Toplevel window with one botton, returns two entry values
    '''
    def __init__(self,
                 title,
                 entry1_label,
                 entry2_label,
                 button_label):
        self.win = tk.Toplevel()
        self.win.title('Find')
        self.head = None
        self.val = None
        self.entry1_label = entry1_label
        self.entry2_label = entry2_label
        self.button_label = button_label
        self.master = tk.Frame(self.win)
        self.master.pack()
        self.upper = tk.Frame(self.master)
        self.upper.pack()
        self.lower = tk.Frame(self.master)
        self.lower.pack()
        self.border = 'grey'
        self.create_widgets()
        
    def create_widgets(self):
        self.entry_frame = tk.Frame(self.upper)
        self.header = tk.Entry(self.entry_frame,
                             highlightthickness = 1,
                             highlightbackground = self.border)
        self.header.grid(row=6, column = 0)
        self.header.focus()
        self.head_lab = tk.Label(self.entry_frame,text = self.entry1_label)
        self.head_lab.grid(row=7, column = 0)
        self.value = tk.Entry(self.entry_frame,
                             highlightthickness = 1,
                             highlightbackground = self.border)
        self.value.grid(row=6, column=4)
        self.val_lab = tk.Label(self.entry_frame, text = self.entry2_label)
        self.val_lab.grid(row=7, column=4)
        self.entry_frame.pack()
        self.button_frame = tk.Frame(self.lower)
        width = max(10, len(self.button_label))
        self.but = tk.Button(self.button_frame,
                          text = self.button_label,
                          width = width,
                          command = self.but_funcs(self.get_both, self.destroy))
        self.but.grid(row=8,column = 10, sticky = tk.E)
        self.but.bind("<Return>", self.but_funcs(self.get_both, self.destroy))   
        self.button_frame.pack()
        self.win.wait_window()

class Triple_entry(Top):
    '''
    Top level window returns three entry field values
    '''

    def __init__(self,
                 title,
                 entry1_label,
                 entry2_label,
                 entry3_label,
                 button_label):
        self.win = tk.Toplevel()
        self.win.geometry('175x200')
        self.win.title(title)
        self.entry1_label = entry1_label
        self.entry2_label = entry2_label
        self.entry3_label = entry3_label
        self.button_label = button_label
        self.master = tk.Frame(self.win)
        self.master.pack()
        self.upper = tk.Frame(self.master)
        self.upper.pack()
        self.lower = tk.Frame(self.master)
        self.lower.pack()
        self.border = 'grey'
        self.create_widgets()
    
        
    def create_widgets(self):
        PADX = 10
        PADY = 10
        
        self.entry_frame = tk.Frame(self.upper)
        self.rolodex = tk.Entry(self.entry_frame,
                             highlightthickness = 1,
                             highlightbackground = self.border)
        self.rolodex.grid(row=6, column = 0, padx = PADX )
        self.rolodex.focus()
        self.rolodex_lab = tk.Label(self.entry_frame,
                                 text = self.entry1_label + '\n')
        self.rolodex_lab.grid(row=7, column = 0, sticky = tk.W, padx = PADX)

        self.header = tk.Entry(self.entry_frame,
                            highlightthickness = 1,
                            highlightbackground = self.border)
        self.header.grid(row=9, column = 0, padx = PADX)
        self.head_lab = tk.Label(self.entry_frame,
                              text = self.entry2_label + '\n')
        self.head_lab.grid(row=10, column = 0, sticky = tk.W, padx = PADX)
        self.value = tk.Entry(self.entry_frame,
                           highlightthickness = 1,
                           highlightbackground = self.border)
        self.value.grid(row=12, column=0, padx = PADX)
        self.val_lab = tk.Label(self.entry_frame,
                             text = self.entry3_label + '\n')
        self.val_lab.grid(row=13, column=0, sticky = tk.W, padx = PADX)
        self.entry_frame.pack()
        self.button_frame = tk.Frame(self.lower)
        width = max(10, len(self.button_label))
        self.but = tk.Button(self.button_frame,
                          text = self.button_label,
                          width = width,
                          command = self.but_funcs(self.get_all, self.destroy))
        self.but.grid(row=15,column = 0, sticky = tk.W)
        self.but.bind("<Return>", self.but_funcs(self.get_all, self.destroy))   
        self.button_frame.pack(side = tk.LEFT)
        self.win.wait_window()
class Single_comboBox(Top):
    '''
    Single Combobox widget, with a single button
    '''
    def __init(self, title, box_values, button_label, box_label = None):
        self.win = tk.Toplevel()
        self.win.geometry('175x200')
        self.win.title(title)
        self.box_values = box_values
        self.button_label = button_label
        self.box_label = box_label
        self.master = Frame(self.win)
        self.master.pacl()
        self.creat_widgets()

    def create_widgets(self):
        box = Combobox(self.master)

class Custom_dialogue(Top):
    '''
    1x1 toplevel window with a Diaplogue box, two buttons, and a checkbox
    '''

    def __init__(self, title, dialogue, check_box_msg):
        self.win = tk.Toplevel()
        self.win.title(title)
        self.msg = dialogue
        self.check_box_msg = check_box_msg
        self.confirm = ''
        self.for_all = ''
        self.master = Frame(self.win)
        self.master.pack()
        self.upper = Frame(self.master)
        self.upper.pack()
        self.lower = Frame(self.master)
        self.lower.pack()
        self.border = 'grey'
        self.create_widgets()

    def create_widgets(self):
        self.msg_frame = Frame(self.upper)
        self.msg = Label(self.msg_frame, text = self.msg)
        self.msg.grid(row = 0, column = 0)
        self.msg_frame.pack()
        self.button_frame = Frame(self.lower)
        self.checkvar = IntVar()
        self.checkvar.set(0)
        self.check = Checkbutton(self.button_frame,
                                 text = self.check_box_msg,
                                 variable = self.checkvar,
                                 onvalue = 1,
                                 offvalue = 0)
        self.yes = Button(self.button_frame,
                          text = 'Yes',
                          command = self.but_funcs(self.yes,self.destroy),
                          width = 10)
        self.no = Button(self.button_frame,
                         text = 'No',
                         command = self.but_funcs(self.no,self.destroy),
                         width = 10)

        self.yes.pack(side = LEFT)
        self.no.pack(side = LEFT)
        self.check.pack(side = BOTTOM)
        self.button_frame.pack()
        self.win.wait_window()


    def yes(self):
        self.confirm = 'yes'
        var = self.checkvar
        if var.get() == 1:
            self.for_all = 'yes'
        if var.get() == 0:
            self.for_all = 'no'

    
    def no(self):
        self.confirm = 'no'
        var = self.checkvar
        if var.get() == 1:
            self.for_all = 'yes'
        if var.get() == 0:
            self.for_all = 'no'




        
class Help_window(Top):
    '''
    A toplevel window, with a listbox on the left and a text display on the
    right.
    '''
    
    def __init__(self, entries, definitions):
        self.entries = entries
        self.definitions = definitions
        self.win = tk.Toplevel()
        self.win.title('Help...')
        self.master = tk.Frame(self.win)
        self.master.rowconfigure(1, weight = 1)
        self.master.columnconfigure(0, weight = 1)
        self.master.columnconfigure(1, weight = 1)
        self.master.pack(fill =tk.BOTH, expand = 1)
        self.text_bar = False
        self.create_widgets()

    def create_widgets(self):
        #list box
        #Frames
        self.label_frame = tk.Frame(self.master)
        self.label_frame.grid(row = 0, column = 0)
        self.label = tk.Label(self.label_frame, text = 'Help with...')
        self.label.pack()
        self.option_frame = tk.Frame(self.master)
        self.option_frame.grid(row = 1, column = 0, sticky = tk.N+tk.E+tk.S+tk.W )
        #Listbox
        self.options = tk.Listbox(self.option_frame)
        for index, option in enumerate(self.entries):
            self.options.insert(index, option)
        self.options.pack(side = tk.LEFT,
                          anchor = tk.NW,
                          fill =tk.BOTH,
                          expand = 1)
        #List Scrollbar
        if len(self.entries) > 10:
            self.list_bar = Scrollbar(self.option_frame,
                                      command = self.options.yview)
            self.options.config(yscrollcommand = self.list_bar.set)
            self.list_bar.pack(side = LEFT,
                               fill = Y)
        #Button
        self.button_frame = tk.Frame(self.master)
        self.button_frame.grid(row = 2, column = 0)
        self.button = tk.Button(self.button_frame,
                             text = 'OK',
                             width = 10,
                             command = self.win.destroy)
        self.button.pack(side = tk.TOP, anchor = tk.N)
        #Keybinding
        self.options.bind('<Double-Button-1>', self.on_double)
        
        
        #Display
        self.display_frame = tk.Frame(self.master)
        self.display_frame.columnconfigure(0, weight = 1)
        self.display_frame.rowconfigure(0, weight = 1)
        self.display_frame.grid(row = 0,
                                rowspan = 3,
                                column = 1,
                                sticky = tk.N+tk.S+tk.E+tk.W)
        self.display = tk.Text(self.display_frame,
                            bg='grey94',
                            wrap = tk.WORD,
                            height = 15,
                            width = 50)
        self.display.pack(side = tk.LEFT, fill= tk.BOTH, expand =1)
        
        

    def on_double(self, event):
        '''
        Return the value selected by user with a double click from listbox
        Display cooresponding Help_text dict entry in text box
        '''
        
        index = self.options.curselection()[0]
        category = self.entries[index]
        text = help_text.display[category]
        self.display.config(state = tk.NORMAL)
        #Clear Text Box of existing text if any
        self.display.delete(1.0,tk.END)
        #Insert desired Help text
        self.display.insert(tk.INSERT,text)
        self.display.config(state = tk.DISABLED)
        lines = int(self.display.index('end-1c').split('.')[0])
        #Vertical scroll bar IF text exceeds n lines
        if lines > 13 and self.text_bar == False:
            self.text_bar = True
            self.bar_frame = tk.Frame(self.display_frame)
            self.bar_frame.pack(side = tk.LEFT, fill = tk.Y)
            self.scrollbar = tk.Scrollbar(self.bar_frame)
            self.scrollbar.pack(side = tk.LEFT, fill =tk.Y)
            self.scrollbar.config(command = self.display.yview)
            self.display.config(yscrollcommand = self.scrollbar.set)
        
        

class Basic_display(Top):
    '''
    1x1 grid, simple canvas to display message with a single button
    '''
    def __init__(self,title, msg):
        self.win = tk.Toplevel()
        self.win.title(title)
        self.master = tk.Frame(self.win)
        self.master.pack()
        self.canvas = tk.Canvas(self.master,
                             borderwidth = 5)
        self.label = tk.Label(self.canvas,
                           text = msg)
        self.button = tk.Button(self.canvas,
                             text = 'Ok',
                             width = 10,
                             borderwidth = 3,
                             command = self.but_funcs(self.destroy))
        self.label.pack()
        self.button.pack()
        self.canvas.pack()


class Read_out(Top):
    '''
    A Toplevel window, used to view the contents of a csv file without the use
    of a third party program. Content of window is unalterable.
    Multi_box() courtesy of python mega widgets.
    '''
    def __init__(self,header,rows,fn):

        #self.win = tk.Toplevel()
        self.header = header
        self.rows = rows
        #rootWin = self.win
        #rootWin.title('Read Out')
        #rootWin.configure(width = 500, height = 300)
        #rootWin.update()
        listBox = Multi_box(fn,labellist = self.header)

        for key,line in rows.items():
            r = {}
            for header in self.header: 
                r[header] = line[header]
            listBox.addrow(r)
        listBox.pack(expand = 1, fill = 'both', padx= 10, pady =10)


#
#  FILE: MCListbox.py
#
#  DESCRIPTION:
#    This file provides a generic Multi-Column Listbox widget.  It is derived
#    from a heavily hacked version of Pmw.ScrolledFrame
#
#  This program is free software; you can redistribute  it and/or modify it
#  under  the terms of  the GNU General  Public License as published by the
#  Free Software Foundation;  either version 2 of the  License, or (at your
#  option) any later version.
#
#  THIS  SOFTWARE  IS PROVIDED   ``AS  IS'' AND   ANY  EXPRESS OR IMPLIED
#  WARRANTIES,   INCLUDING, BUT NOT  LIMITED  TO, THE IMPLIED WARRANTIES OF
#  MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN
#  NO  EVENT  SHALL   THE AUTHOR  BE    LIABLE FOR ANY   DIRECT, INDIRECT,
#  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
#  NOT LIMITED   TO, PROCUREMENT OF  SUBSTITUTE GOODS  OR SERVICES; LOSS OF
#  USE, DATA,  OR PROFITS; OR  BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
#  ANY THEORY OF LIABILITY, WHETHER IN  CONTRACT, STRICT LIABILITY, OR TORT
#  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
#  THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
#  You should have received a copy of the  GNU General Public License along
#  with this program; if not, write  to the Free Software Foundation, Inc.,
#  675 Mass Ave, Cambridge, MA 02139, USA.
#


class Multi_box(Pmw.MegaWidget):
    def __init__(self,fn,parent = None, **kw):
        colors = Pmw.Color.getdefaultpalette(parent)
        parent = tk.Toplevel()
        parent.title('{}'.format(fn))
        # Define the megawidget options.
        INITOPT = Pmw.INITOPT
        optiondefs = (
            #('borderframe',      1,                          INITOPT),
            ('horizflex',        'fixed',                    self._horizflex),
            ('horizfraction',    0.05,                       INITOPT),
            ('hscrollmode',      'dynamic',                  self._hscrollMode),
            ('labelmargin',      0,                          INITOPT),
            ('labelpos',         None,                       INITOPT),
            ('scrollmargin',     2,                          INITOPT),
            ('usehullsize',      0,                          INITOPT),
            ('vertflex',         'fixed',                    self._vertflex),
            ('vertfraction',     0.05,                       INITOPT),
            ('vscrollmode',      'dynamic',                  self._vscrollMode),
            ('labellist',        None,                       INITOPT),
            ('selectbackground', colors['selectBackground'], INITOPT),
            ('selectforeground', colors['selectForeground'], INITOPT),
            ('background',       colors['background'],       INITOPT),
            ('foreground',       colors['foreground'],       INITOPT),
            ('command',          None,                       None),
            ('dblclickcommand',  None,                       None),
        )
        self.defineoptions(kw, optiondefs)

        # Initialise the base class (after defining the options).
        Pmw.MegaWidget.__init__(self, parent)

        self._numcolumns = len(self['labellist'])
        self._columnlabels = self['labellist']
        self._lineid = 0
        self._numrows = 0
        self._lineitemframes = []
        self._lineitems = []
        self._lineitemdata = {}
        self._labelframe = {}
        self._cursel = []

        # Create the components.
        self.origInterior = Pmw.MegaWidget.interior(self)

        if self['usehullsize']:
            self.origInterior.grid_propagate(0)

        # Create a frame widget to act as the border of the clipper.
        self._borderframe = self.createcomponent('borderframe',
                                                 (), None,
                                                 tk.Frame,
                                                 (self.origInterior,),
                                                 relief = 'sunken',
                                                 borderwidth = 2,
                                                 )
        self._borderframe.grid(row = 2, column = 2,
                               rowspan = 2, sticky = 'news')

        # Create the clipping windows.
        self._hclipper = self.createcomponent('hclipper',
                                              (), None,
                                              tk.Frame,
                                              (self._borderframe,),
                                              width = 400,
                                              height = 300,
                                              )
        self._hclipper.pack(fill = 'both', expand = 1)

        self._hsframe = self.createcomponent('hsframe', (), None,
                                             tk.Frame,
                                             (self._hclipper,),
                                             )


        self._vclipper = self.createcomponent('vclipper',
                                              (), None,
                                              tk.Frame,
                                              (self._hsframe,),
                                              #width = 400,
                                              #height = 300,
                                              highlightthickness = 0,
                                              borderwidth = 0,
                                              )

        self._vclipper.grid(row = 1, column = 0,
                            columnspan = self._numcolumns,
                            sticky = 'news')#, expand = 1)
        self._hsframe.grid_rowconfigure(1, weight = 1)#, minsize = 300)


        gridcolumn = 0
        for labeltext in self._columnlabels:
            lframe = self.createcomponent(labeltext+'frame', (), None,
                                          tk.Frame,
                                          (self._hsframe,),
                                          borderwidth = 1,
                                          relief = 'raised',
                                          )
            label = self.createcomponent(labeltext, (), None,
                                         tk.Label,
                                         (lframe,),
                                         text = labeltext,
                                         )
            label.pack(expand = 0, fill = 'y', side = 'left')
            lframe.grid(row = 0, column = gridcolumn, sticky = 'ews')
            self._labelframe[labeltext] = lframe
            #lframe.update()
            #print lframe.winfo_reqwidth()
            self._hsframe.grid_columnconfigure(gridcolumn, weight = 1)
            gridcolumn = gridcolumn + 1

        lframe.update()
        self._labelheight = lframe.winfo_reqheight()
        self.origInterior.grid_rowconfigure(2, minsize = self._labelheight + 2)

        self.origInterior.grid_rowconfigure(3, weight = 1, minsize = 0)
        self.origInterior.grid_columnconfigure(2, weight = 1, minsize = 0)

        # Create the horizontal scrollbar
        self._horizScrollbar = self.createcomponent('horizscrollbar',
                                                    (), 'Scrollbar',
                                                    tk.Scrollbar,
                                                    (self.origInterior,),
                                                    orient='horizontal',
                                                    command=self._xview
                                                    )

        # Create the vertical scrollbar
        self._vertScrollbar = self.createcomponent('vertscrollbar',
                                                   (), 'Scrollbar',
                                                   tk.Scrollbar,
                                                   (self.origInterior,),
                                                   #(self._hclipper,),
                                                   orient='vertical',
                                                   command=self._yview
                                                   )

        self.createlabel(self.origInterior, childCols = 3, childRows = 4)

        # Initialise instance variables.
        self._horizScrollbarOn = 0
        self._vertScrollbarOn = 0
        self.scrollTimer = None
        self._scrollRecurse = 0
        self._horizScrollbarNeeded = 0
        self._vertScrollbarNeeded = 0
        self.startX = 0
        self.startY = 0
        self._flexoptions = ('fixed', 'expand', 'shrink', 'elastic')

        # Create a frame in the clipper to contain the widgets to be
        # scrolled.
        self._vsframe = self.createcomponent('vsframe',
                                             (), None,
                                             tk.Frame,
                                             (self._vclipper,),
                                             #height = 300,
                                             #borderwidth = 4,
                                             #relief = 'groove',
                                             )

        # Whenever the clipping window or scrolled frame change size,
        # update the scrollbars.
        self._hsframe.bind('<Configure>', self._reposition)
        self._vsframe.bind('<Configure>', self._reposition)
        self._hclipper.bind('<Configure>', self._reposition)
        self._vclipper.bind('<Configure>', self._reposition)

        #elf._vsframe.bind('<Button-1>', self._vsframeselect)

        # Check keywords and initialise options.
        self.initialiseoptions()

    def destroy(self):
        if self.scrollTimer is not None:
            self.after_cancel(self.scrollTimer)
            self.scrollTimer = None
        Pmw.MegaWidget.destroy(self)

    # ======================================================================

    # Public methods.

    def interior(self):
        return self._vsframe

    # Set timer to call real reposition method, so that it is not
    # called multiple times when many things are reconfigured at the
    # same time.
    def reposition(self):
        if self.scrollTimer is None:
            self.scrollTimer = self.after_idle(self._scrollBothNow)



    def insertrow(self, index, rowdata):
        #if len(rowdata) != self._numcolumns:
        #    raise ValueError, 'Number of items in rowdata does not match number of columns.'
        if index > self._numrows:
            index = self._numrows

        rowframes = {}
        for columnlabel in self._columnlabels:
            celldata = rowdata.get(columnlabel)
            cellframe = self.createcomponent(('cellframeid.%d.%s'%(self._lineid,
                                                                   columnlabel)),
                                             (), ('Cellframerowid.%d'%self._lineid),
                                             tk.Frame,
                                             (self._vsframe,),
                                             background = self['background'],
                                             #borderwidth = 1,
                                             #relief = 'flat'
                                             )

            cellframe.bind('<Double-Button-1>', self._cellframedblclick)
            cellframe.bind('<Button-1>', self._cellframeselect)

            if celldata:
                cell = self.createcomponent(('cellid.%d.%s'%(self._lineid,
                                                             columnlabel)),
                                            (), ('Cellrowid.%d'%self._lineid),
                                            tk.Label,
                                            (cellframe,),
                                            background = self['background'],
                                            foreground = self['foreground'],
                                            text = celldata,
                                            )

                cell.bind('<Double-Button-1>', self._celldblclick)
                cell.bind('<Button-1>', self._cellselect)

                cell.pack(expand = 0, fill = 'y', side = 'left', padx = 1, pady = 1)
            rowframes[columnlabel] = cellframe

        self._lineitemdata[self._lineid] = rowdata
        self._lineitems.insert(index, self._lineid)
        self._lineitemframes.insert(index, rowframes)
        self._numrows = self._numrows + 1
        self._lineid = self._lineid + 1

        self._placedata(index)

    def _placedata(self, index = 0):
        gridy = index
        for rowframes in self._lineitemframes[index:]:
            gridx = 0
            for columnlabel in self._columnlabels:
                rowframes[columnlabel].grid(row = gridy,
                                           column = gridx,
                                           sticky = 'news')
                gridx = gridx + 1
            gridy = gridy + 1



    def addrow(self, rowdata):
        self.insertrow(self._numrows, rowdata)

    def delrow(self, index):
        rowframes = self._lineitemframes.pop(index)
        for columnlabel in self._columnlabels:
            rowframes[columnlabel].destroy()
        self._placedata(index)
        self._numrows = self._numrows - 1
        del self._lineitems[index]
        if index in self._cursel:
            self._cursel.remove(index)


    def curselection(self):
        # Return a tuple of just one element as this will probably be the
        # interface used in a future implementation when multiple rows can
        # be selected at once.
        return tuple(self._cursel)

    def getcurselection(self):
        # Return a tuple of just one row as this will probably be the
        # interface used in a future implementation when multiple rows can
        # be selected at once.
        sellist = []
        for sel in self._cursel:
            sellist.append(self._lineitemdata[self._lineitems[sel]])
        return tuple(sellist)

    # ======================================================================

    # Configuration methods.

    def _hscrollMode(self):
        # The horizontal scroll mode has been configured.

        mode = self['hscrollmode']


        if mode == 'static':
            if not self._horizScrollbarOn:
                self._toggleHorizScrollbar()
        elif mode == 'dynamic':
            if self._horizScrollbarNeeded != self._horizScrollbarOn:
                self._toggleHorizScrollbar()
        elif mode == 'none':
            if self._horizScrollbarOn:
                self._toggleHorizScrollbar()
        else:
            message = 'bad hscrollmode option "%s": should be static, dynamic, or none' % mode
            raise ValueError(message)

    def _vscrollMode(self):
        # The vertical scroll mode has been configured.

        mode = self['vscrollmode']

        if mode == 'static':
            if not self._vertScrollbarOn:
                self._toggleVertScrollbar()
        elif mode == 'dynamic':
            if self._vertScrollbarNeeded != self._vertScrollbarOn:
                self._toggleVertScrollbar()
        elif mode == 'none':
            if self._vertScrollbarOn:
                self._toggleVertScrollbar()
        else:
            message = 'bad vscrollmode option "%s": should be static, dynamic, or none' % mode
            raise ValueError(message)

    def _horizflex(self):
        # The horizontal flex mode has been configured.

        flex = self['horizflex']

        if flex not in self._flexoptions:
            message = 'bad horizflex option "%s": should be one of %s' % \
                    mode, str(self._flexoptions)
            raise ValueError(message)

        self.reposition()

    def _vertflex(self):
        # The vertical flex mode has been configured.

        flex = self['vertflex']

        if flex not in self._flexoptions:
            message = 'bad vertflex option "%s": should be one of %s' % \
                    mode, str(self._flexoptions)
            raise ValueError(message)

        self.reposition()



    # ======================================================================

    # Private methods.

    def _reposition(self, event):
        gridx = 0
        for col in self._columnlabels:
            maxwidth = self._labelframe[col].winfo_reqwidth()
            for row in self._lineitemframes:
                cellwidth = row[col].winfo_reqwidth()
                if cellwidth > maxwidth:
                    maxwidth = cellwidth
            self._hsframe.grid_columnconfigure(gridx, minsize = maxwidth)
            gridwidth = self._hsframe.grid_bbox(column = gridx, row = 0)[2]
            if self['horizflex'] in ('expand', 'elastic') and gridwidth > maxwidth:
                maxwidth = gridwidth
            self._vsframe.grid_columnconfigure(gridx, minsize = maxwidth)
            gridx = gridx + 1



        self._vclipper.configure(height = self._hclipper.winfo_height() - self._labelheight)

        self.reposition()

    # Called when the user clicks in the horizontal scrollbar.
    # Calculates new position of frame then calls reposition() to
    # update the frame and the scrollbar.
    def _xview(self, mode, value, units = None):

        if mode == 'moveto':
            frameWidth = self._hsframe.winfo_reqwidth()
            self.startX = float(value) * float(frameWidth)
        else:
            clipperWidth = self._hclipper.winfo_width()
            if units == 'units':
                jump = int(clipperWidth * self['horizfraction'])
            else:
                jump = clipperWidth

            if value == '1':
                self.startX = self.startX + jump
            else:
                self.startX = self.startX - jump

        self.reposition()

    # Called when the user clicks in the vertical scrollbar.
    # Calculates new position of frame then calls reposition() to
    # update the frame and the scrollbar.
    def _yview(self, mode, value, units = None):

        if mode == 'moveto':
            frameHeight = self._vsframe.winfo_reqheight()
            self.startY = float(value) * float(frameHeight)
        else:
            clipperHeight = self._vclipper.winfo_height()
            if units == 'units':
                jump = int(clipperHeight * self['vertfraction'])
            else:
                jump = clipperHeight

            if value == '1':
                self.startY = self.startY + jump
            else:
                self.startY = self.startY - jump

        self.reposition()

    def _getxview(self):

        # Horizontal dimension.
        clipperWidth = self._hclipper.winfo_width()
        frameWidth = self._hsframe.winfo_reqwidth()
        if frameWidth <= clipperWidth:
            # The scrolled frame is smaller than the clipping window.

            self.startX = 0
            endScrollX = 1.0

            if self['horizflex'] in ('expand', 'elastic'):
                relwidth = 1
            else:
                relwidth = ''
        else:
            # The scrolled frame is larger than the clipping window.

            if self['horizflex'] in ('shrink', 'elastic'):
                self.startX = 0
                endScrollX = 1.0
                relwidth = 1
            else:
                if self.startX + clipperWidth > frameWidth:
                    self.startX = frameWidth - clipperWidth
                    endScrollX = 1.0
                else:
                    if self.startX < 0:
                        self.startX = 0
                    endScrollX = (self.startX + clipperWidth) / float(frameWidth)
                relwidth = ''

        # Position frame relative to clipper.
        self._hsframe.place(x = -self.startX, relwidth = relwidth)
        return (self.startX / float(frameWidth), endScrollX)

    def _getyview(self):

        # Vertical dimension.
        clipperHeight = self._vclipper.winfo_height()
        frameHeight = self._vsframe.winfo_reqheight()
        if frameHeight <= clipperHeight:
            # The scrolled frame is smaller than the clipping window.

            self.startY = 0
            endScrollY = 1.0

            if self['vertflex'] in ('expand', 'elastic'):
                relheight = 1
            else:
                relheight = ''
        else:
            # The scrolled frame is larger than the clipping window.

            if self['vertflex'] in ('shrink', 'elastic'):
                self.startY = 0
                endScrollY = 1.0
                relheight = 1
            else:
                if self.startY + clipperHeight > frameHeight:
                    self.startY = frameHeight - clipperHeight
                    endScrollY = 1.0
                else:
                    if self.startY < 0:
                        self.startY = 0
                    endScrollY = (self.startY + clipperHeight) / float(frameHeight)
                relheight = ''

        # Position frame relative to clipper.
        self._vsframe.place(y = -self.startY, relheight = relheight)
        return (self.startY / float(frameHeight), endScrollY)

    # According to the relative geometries of the frame and the
    # clipper, reposition the frame within the clipper and reset the
    # scrollbars.
    def _scrollBothNow(self):
        self.scrollTimer = None

        # Call update_idletasks to make sure that the containing frame
        # has been resized before we attempt to set the scrollbars.
        # Otherwise the scrollbars may be mapped/unmapped continuously.
        self._scrollRecurse = self._scrollRecurse + 1
        self.update_idletasks()
        self._scrollRecurse = self._scrollRecurse - 1
        if self._scrollRecurse != 0:
            return

        xview = self._getxview()
        yview = self._getyview()
        self._horizScrollbar.set(xview[0], xview[1])
        self._vertScrollbar.set(yview[0], yview[1])

        self._horizScrollbarNeeded = (xview != (0.0, 1.0))
        self._vertScrollbarNeeded = (yview != (0.0, 1.0))

        # If both horizontal and vertical scrollmodes are dynamic and
        # currently only one scrollbar is mapped and both should be
        # toggled, then unmap the mapped scrollbar.  This prevents a
        # continuous mapping and unmapping of the scrollbars.
        if (self['hscrollmode'] == self['vscrollmode'] == 'dynamic' and
                self._horizScrollbarNeeded != self._horizScrollbarOn and
                self._vertScrollbarNeeded != self._vertScrollbarOn and
                self._vertScrollbarOn != self._horizScrollbarOn):
            if self._horizScrollbarOn:
                self._toggleHorizScrollbar()
            else:
                self._toggleVertScrollbar()
            return

        if self['hscrollmode'] == 'dynamic':
            if self._horizScrollbarNeeded != self._horizScrollbarOn:
                self._toggleHorizScrollbar()

        if self['vscrollmode'] == 'dynamic':
            if self._vertScrollbarNeeded != self._vertScrollbarOn:
                self._toggleVertScrollbar()

    def _toggleHorizScrollbar(self):

        self._horizScrollbarOn = not self._horizScrollbarOn

        interior = self.origInterior
        if self._horizScrollbarOn:
            self._horizScrollbar.grid(row = 5, column = 2, sticky = 'news')
            interior.grid_rowconfigure(4, minsize = self['scrollmargin'])
        else:
            self._horizScrollbar.grid_forget()
            interior.grid_rowconfigure(4, minsize = 0)

    def _toggleVertScrollbar(self):

        self._vertScrollbarOn = not self._vertScrollbarOn

        interior = self.origInterior
        if self._vertScrollbarOn:
            self._vertScrollbar.grid(row = 3, column = 4, sticky = 'news')
            interior.grid_columnconfigure(3, minsize = self['scrollmargin'])
        else:
            self._vertScrollbar.grid_forget()
            interior.grid_columnconfigure(3, minsize = 0)

    # ======================================================================

    # Selection methods.

    #def _vsframeselect(self, event):
    #    print 'vsframe event x: %d  y: %d'%(event.x, event.y)
    #    col, row = self._vsframe.grid_location(event.x, event.y)
    #    self._select(col, row)

    def _cellframeselect(self, event):
        #print 'cellframe event x: %d  y: %d'%(event.x, event.y)
        x = event.widget.winfo_x()
        y = event.widget.winfo_y()
        #col, row = self._vsframe.grid_location(x + event.x, y + event.y)
        self._select(x + event.x, y + event.y)#(col, row)

    def _cellselect(self, event):
        #print 'cell event x: %d  y: %d'%(event.x, event.y)
        lx = event.widget.winfo_x()
        ly = event.widget.winfo_y()
        parent = event.widget.pack_info()['in']
        fx = parent.winfo_x()
        fy = parent.winfo_y()
        #col, row = self._vsframe.grid_location(fx + lx + event.x, fy + ly + event.y)
        self._select(fx + lx + event.x, fy + ly + event.y)#(col, row)

    def _select(self, x, y):
        col, row = self._vsframe.grid_location(x, y)
        #print 'Clicked on col: %d  row: %d'%(col,row)
        cfg = {}
        lineid = self._lineitems[row]
        cfg['Cellrowid.%d_foreground'%lineid] = self['selectforeground']
        cfg['Cellrowid.%d_background'%lineid] = self['selectbackground']
        cfg['Cellframerowid.%d_background'%lineid] = self['selectbackground']
        #cfg['Cellframerowid%d_relief'%row] = 'raised'

        if self._cursel != []:
            cursel = self._cursel[0]
            lineid = self._lineitems[cursel]
            if cursel != None and cursel != row:
                cfg['Cellrowid.%d_foreground'%lineid] = self['foreground']
                cfg['Cellrowid.%d_background'%lineid] = self['background']
                cfg['Cellframerowid.%d_background'%lineid] = self['background']
                #cfg['Cellframerowid%d_relief'%cursel] = 'flat'

        self.configure(*(), **cfg)
        self._cursel = [row]

        cmd = self['command']
        if isinstance(cmd, collections.Callable):
            cmd()



    def _cellframedblclick(self, event):
        #print 'double click cell frame'
        cmd = self['dblclickcommand']
        if isinstance(cmd, collections.Callable):
            cmd()

    def _celldblclick(self, event):
        #print 'double click cell'
        cmd = self['dblclickcommand']
        if isinstance(cmd, collections.Callable):
            cmd()

if __name__ == '__main__':

    Double_entry('Title', 'Entry1', 'Entry2', 'button')
    '''
    rootWin = tk.Toplevel()

    Pmw.initialise()

    rootWin.title('MultiColumnListbox Demo')
    rootWin.configure(width = 500, height = 300)
    rootWin.update()

    def dbl():
        print(listbox.getcurselection())

    listbox = Read_out(rootWin
                                           )

    
    #print 'start adding item'
    for i in range(20):
        r = {}
        for j in range(5):
            r[('Column %d'%j)] = 'Really long item name %d'%i
        listbox.addrow(r)
    #print 'items added'

    listbox.pack(expand = 1, fill = 'both', padx = 10, pady = 10)


    exitButton = tk.Button(rootWin, text="Quit", command=rootWin.quit)
    exitButton.pack(side = 'left', padx = 10, pady = 10)

    rootWin.mainloop()
    '''

        
       
        






