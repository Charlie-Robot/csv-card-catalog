import csv
import os
import time
import pickle
from top_wins import *
from tkinter.messagebox import showerror, askquestion, showwarning

class Spreadsheet:
    def __init__(self, headers = None, fn = None, restval = None):
        '''
        A CSV file object:
        Designed to work in conjuction with a specific set of data delivered by
        a GUI interface which collects Union Membership card input.
        File name defaults to today's date if no file name specified. If file
        already exists, name appended with (n): 12.10.2013(1).csv.
        Takes a dict of dicts as input: {rolodex: {data row}}
        *Rolodex is a unique ID primarliy updated by GUI, and stored with pickle
        '''

        #File name defaults to todays date FN = '01.01.2016.csv'
        DATE_FORMAT = '%m/%d/%Y'
        DEFAULT_EXT = '.csv'
        self.DATE = time.strftime(DATE_FORMAT) 
        self.FN_DATE = self.DATE.replace('/', '.')
        self.FN = fn
        self.ROLODEX_FN = 'rolodex.pkl'

        #List of Header Field Names
        self.HEADER = headers
        
        #read/write dicts used to parse and process data rows
        self.READ_DICT =  {}
        self.WRITE_DICT = {}

        #Rolodex starting index, rolodex number is unique to data row
        self.ROLODEX = 100
        
        #Value to fill blank cells
        self.RESTVAL = restval
        if self.RESTVAL == None:
            self.RESTVAL = '-'
        
        #In case of duplicate file names, add (n) to file name: duplicate(1).csv    
        self.FN_PADDING = 0
        if self.FN == None:
            self.FN = self.FN_DATE + DEFAULT_EXT
            while os.path.isfile(self.FN):
                self.FN_PADDING += 1
                pad  = '(' + str(self.FN_PADDING) + ')'
                #sheet.FN default is dd.mm.yyyy
                new_fn = self.FN_DATE + pad + '.csv'
                self.FN = new_fn
         
    def __str__(self):
        return "Spreadsheet Object\nFile Name: {}\nRolodex: {}\nDirectory: {}"\
               .format(
            self.FN,self.ROLODEX, os.getcwd())
    
    def __repr__(self):
        return Spreadsheet(headers = self.HEADER, fn = 'Representing')

    #Public Methods
    def read(self, fn = None):
        '''
        Reads rows of lists into self.READ_DICT from self.FN and convert/store
        them in self.WRITE_DICT as {Rolodex ID: {headers:values}},
        update self.ROLODEX to highest known value
        '''
        if fn == None:
            fn = self.FN
        try:
            with open(fn, 'r', newline = '') as content:
                csv_reader = csv.reader(content)
                
                
                for row in csv_reader:
                    _ID = row[11]
                    # _ID == 'Rolodex', indicates that row is Header Values
                    if _ID == 'Rolodex':
                        pass
                    else:
                        self.READ_DICT[_ID] = row
                    
                for key, row in self.READ_DICT.items():
                    self.WRITE_DICT[key] = {header:value for
                                            header, value in
                                            zip(self.HEADER, row)}
        except FileNotFoundError:
            print("No File Specified")
            pass
        self.ROLODEX = max(int(max(self.READ_DICT.keys())), self.get_dex())      

    def write(self, entries, mode='a', newline = '',
              encoding ='utf8', dialect = None):
        '''
        entries is a dict of dict's with rolodex# for unique ID and keys for
        outer dict, inner dict is header:entry paired values for rows in sheet.
        Default mode == 'a' to append new rows of data,
        if mode == 'w', HEADER is written before new rows. 
        
        '''
        if entries == None:
            entries = {}
        #Remove key:value pairs if value == '',
        #replaced with key:restval when written
        for row in entries.values():
            blanks = []
            for header, value in row.items():
                if value == '':
                    blanks.append(header)
            for header in blanks:
                row.pop(header)
        with open(self.FN, mode, newline='') as csvfile:
            writer = csv.DictWriter(csvfile,
                                    dialect = dialect,
                                    fieldnames = self.HEADER,
                                    restval = self.RESTVAL)
            if mode == 'w':
                writer.writeheader()
            for key in sorted(entries.keys()):
                writer.writerow(entries[key])
           
    def load(self, fn):
        """
        Clears all entries currently in READ_DICT and WRITE_DICT,
        set FN attribute to new file name and read in new data.  
        """
        self.READ_DICT = {}
        self.WRITE_DICT = {}
        if os.path.isfile(fn):
            self.FN = fn
        #Read in fresh data; raises ValueError if empty sheet loaded
        try:
            self.read()
        except ValueError:
            print('Read Empty Sheet')
            pass
        
    def find(self, header, value):
        '''
        Isolate any cells with a given value under a specfic fieldname,
        collect all rows which meet specified criterion, overwrite
        self.WRITE_DICT with found data rows if any.
        '''
        found = {}
        header = header.title()
        try:
            for indx, row in self.WRITE_DICT.items():
                if row[header] == value:
                    found[indx] = row
            if len(found) > 0:
                self.WRITE_DICT = found
                msg = "{} cards found where '{}' = '{}'".format(
                                                      len(self.WRITE_DICT),
                                                      header,
                                                      value)
            elif len(found) == 0:
                msg = "Searched: '{}' = '{}'\nNo matches where found".format(header, value)
        except KeyError:
            msg = "Invalid Header: '{}'".format(header)
        return msg
       
    def remove_row(self,*args):
        '''
        Remove an entry or entries from self.WRITE_DICT based on Rolodex ID
        and/or header == value, ID is a string of comma seperated unique
        ID's or range of ID's,
        example: '101,102,200-250' -OR- '101' -OR- '200-250'
        Value is a string value of any given cell
        '''
        keys = []
        dic = self.WRITE_DICT
        rolodex, header, value = args
        header = header.title()
        row_ID = dic.keys()
        #Remove ALL rows if Rolodex where header == value
        if rolodex != '' and value != '' and header != '':
            for ID in rolodex.split(','):
                if ID.strip() in row_ID:
                    if dic[ID][header] == value:
                        keys.append(ID)
                else:
                    #Remove a range of Rolodex numbers
                    if '-' in ID:
                        indx = ID.index('-')
                        start = ID[:indx].strip()
                        end = ID[indx + 1:].strip()
                        for n in range(int(start), int(end) + 1):
                            ID = str(n)
                            if ID in dic.keys():
                                if dic[ID][header] == value:
                                    keys.append(ID)
        #Remove ALL rows where header == value regardless of Rolodex
        if rolodex == '' and value != '' and header != '':
            for ID, row in dic.items():
                if row[header] == value:
                   keys.append(ID)
        #Remove ALL rows where Rolodex in self.WRITE_DICT.keys()
        if rolodex != '' and  value == '' or header == '':
            for ID in rolodex.split(','):
                if ID.strip() in dic.keys():
                    keys.append(ID)
                else:
                    #Remove a range of Rolodex numbers
                    if '-' in ID:
                        indx = ID.index('-')
                        start = ID[:indx].strip()
                        end = ID[indx + 1:].strip()
                        for n in range(int(start), int(end) + 1):
                            ID = str(n)
                            if ID in dic.keys():
                                keys.append(ID)
        for ID in keys:
            dic.pop(ID.strip())
        
        return keys
            
    def remove_col(self, header):   
        '''
        Purge read-in rows from READ_DICT and removes all data under a given
        header from every dict in WRITE_DICT
        '''
        self.READ_DICT = {}
        for indx, row in self.WRITE_DICT.items():
            row[header] = ''

    def edit_cell(self, ID, header, new_val, over_ride = False):
        '''
        Alter all values under a sinlge header in a given row or list of rows,
        as identified by their Rolodex ID's. data_row['Rolodex'] not alterable
        unless over_ride = True
        '''
        keys = []
        for key in ID.split(','):
            if '-' in key:
                indx = key.index('-')
                start = key[:indx].strip()
                end = key[indx + 1:].strip()
                for n in range(int(start), int(end) + 1):
                    key = str(n)
                    if key in self.WRITE_DICT.keys():
                        keys.append(key)
            elif key.strip() in self.WRITE_DICT.keys():
                keys.append(key.strip())
            
        if  header == 'Rolodex' and over_ride:
            for rolodex in ID.split(','):
                if rolodex in self.WRITE_DICT.keys():
                    self.WRITE_DICT[ID].pop('Rolodex')
                    self.WRITE_DICT[ID]['Rolodex'] = new_val
                    row = self.WRITE_DICT.pop(ID)
                    self.WRITE_DICT[new_val] = row
                       
        for rolodex in keys:
            if rolodex in self.WRITE_DICT.keys() and header != 'Rolodex':
                self.WRITE_DICT[rolodex][header] = new_val
            

    def merge(self, fn2, **kwargs):
        '''
        An existing CSV file is combined with current WRITE_DICT and written to
        a new file.  In event that duplicate Rolodex values exist, top_level GUI
        prompts to change Rolodex, or overwrite duplicate ID with option to do
        same action for all duplicate ID's that remain
        '''
        #Top level copy of self.WRITE_DICT
        replace = {}
        
        replace.update(self.WRITE_DICT)
        
        conditions = {'TESTING' : False,
                      'CHANGE_YES_ALL' : False,
                      'CHANGE_YES_ONE' : False,
                      'CHANGE_NO_ALL' : False,
                      'CHANGE_NO_ONE' : False,
                      'OVERWRITE_YES_ALL' : False,
                      'OVERWRITE_YES_ONE' : False,
                      'OVERWRITE_NO_ALL' : False,
                      'OVERWRITE_NO_ONE' : False,}
        
        for key, value in kwargs.items():
            conditions[key] = value
        
        TESTING = conditions['TESTING']
        CHANGE_YES_ALL = conditions['CHANGE_YES_ALL']
        CHANGE_YES_ONE = conditions['CHANGE_YES_ONE']
        CHANGE_NO_ALL = conditions['CHANGE_NO_ALL']
        CHANGE_NO_ONE = conditions['CHANGE_NO_ONE']
        OVERWRITE_YES_ALL = conditions['OVERWRITE_YES_ALL']
        OVERWRITE_YES_ONE = conditions['OVERWRITE_YES_ONE']
        OVERWRITE_NO_ALL = conditions['OVERWRITE_NO_ALL']
        OVERWRITE_NO_ONE = conditions['OVERWRITE_NO_ONE']
       

        fn = self.FN
        new, ext = os.path.splitext(fn)
        path, FN2 = os.path.split(fn2)
        path2 = os.path.join(path, fn2)
        merged_fn = new + '_merged' + ext
        merged = Spreadsheet(self.HEADER, fn = merged_fn)
        msg = ''
        if os.path.isfile(path2) and os.path.isfile(fn):
            merged.load(fn2)
                          
            for key in sorted(merged.WRITE_DICT.keys()):
                if key in self.WRITE_DICT.keys():
                    
                    #If duplicate Rolodex Found, prompt for change/overwrite/ignore
                    if CHANGE_YES_ALL == False and CHANGE_NO_ALL == False and\
                       not OVERWRITE_YES_ALL and not OVERWRITE_NO_ALL:
                        msg = 'Duplicate Rolodex Found: {}\nChange {} Rolododex to: {}?'.format(
                            key, FN2,  self.get_dex())
                        if not TESTING:
                            #supress GUI for testing
                            change = Custom_dialogue('Warning', msg, 'Do this for all files?')

                            if change.confirm == 'yes' and change.for_all == 'yes':
                                CHANGE_YES_ALL = True
                                CHANGE_NO_ALL = False
                            if change.confirm == 'yes' and change.for_all == 'no':
                                CHANGE_YES_ONE = True
                            if change.confirm == 'no' and change.for_all == 'yes':
                                CHANGE_YES_ALL = False
                                CHANGE_NO_ALL = True
                            if change.confirm == 'no' and change.for_all == 'no':
                                CHANGE_NO_ONE = True
                    
                    if CHANGE_YES_ALL or CHANGE_YES_ONE:
                        #Update duplicate Rolodex(s) to highest known Rolodex value
                        new_ID = str(self.get_dex() + 1)
                        merged.WRITE_DICT[new_ID] =\
                        merged.WRITE_DICT.pop(key)
                        merged.WRITE_DICT[new_ID]['Rolodex'] = new_ID
                        #Update working Spreadsheet rolodex
                        self.ROLODEX += 1
                        CHANGE_YES_ONE = False
                        #continue
                    
                    #If not changing Rolodex: Overwrite Yes/No

                    if CHANGE_NO_ALL or CHANGE_NO_ONE:
                        WARNING = False
                        CHANGE_NO_ONE = False
                        CHANGE_YES_ONE = False
                        msg = '\nOverwrite {} in {} with {} from {}?\n'.format(
                            key, os.path.basename(self.FN), key, os.path.basename(fn2))
                        if not TESTING and not OVERWRITE_YES_ALL and not OVERWRITE_NO_ALL:
                            #supress GUI for testing
                            overwrite = Custom_dialogue("Overwrite", msg, 'Do this for all files?')
                            if overwrite.confirm == 'yes' and overwrite.for_all == 'yes':
                                OVERWRITE_YES_ALL = True
                                WARNING = True
                                msg = 'Remaining files from {} will not be added to {}'.format(
                                       os.path.basename(fn),
                                       os.path.basename(merged_fn))
                            if overwrite.confirm == 'yes' and overwrite.for_all == 'no':
                                OVERWRITE_YES_ONE = True
                                WARNING = True
                                msg = "{} from {} will not be added to {}".format(
                                       key,
                                       os.path.basename(self.FN),
                                       os.path.basename(merged_fn))
                            if overwrite.confirm == 'no' and overwrite.for_all == 'yes':
                                OVERWRITE_NO_ALL = True
                                WARNING = True
                                msg = 'Remaining files from {} will not be added to {}'.format(
                                       os.path.basename(fn2),
                                       os.path.basename(merged_fn))
                            if overwrite.confirm == 'no' and overwrite.for_all == 'no':
                                OVERWRITE_NO_ONE = True
                                WARNING = True
                                msg = "{} from {} will not be add to {}".format(
                                       key,
                                       os.path.basename(merged.FN),
                                       os.path.basename(merged_fn))
               
                        if WARNING:
                            showwarning(title = 'WARNING!', message = msg)
                            WARNING = False
                      
                   
                    if OVERWRITE_YES_ALL or OVERWRITE_YES_ONE:
                        replace[key] = merged.WRITE_DICT[key]
                        OVERWRITE_YES_ONE = False
                        continue

            
            merged.WRITE_DICT.update(replace)
            
            
            merged.FN = merged_fn
            merged.write(merged.WRITE_DICT)
            msg = 'Files Merged: {0} with {1}\n\n Merged File: {2}'.format(
                   fn, fn2, merged_fn)
            
            return msg                        
                                                   
    def read_out(self):
        '''
        A generator object to feed rows of currently open CSV file to a top level
        window that will visually represent them without a third party program.
        '''
        
         
        for row in self.WRITE_DICT.values():
            yield row

    #Private Methods
            
    def get_dex(self):
        '''
        Check for pickled rolodex file and returns highest known rolodex as int 
        '''
        if os.path.isfile('rolodex.pkl'):
            with open('rolodex.pkl', 'rb') as n:
                dex = pickle.load(n)
                dex = int(dex)
                rolodex = max(dex, self.ROLODEX)
        else:
            rolodex = self.ROLODEX
            
        return rolodex



if __name__ == '__main__':

    sheet = Spreadsheet()
    print(sheet)

