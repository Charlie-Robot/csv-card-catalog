import csv
import os

class Spreadsheet:
    def __init__(self, headers = None, fn = None):
        '''
        Inialize with a dict of nested dicts: row_dict[row_ID] = {header:data};
        header and fn are a list and a string, header is list of headers to use,
        fn is the name of the spreadsheet file attached to the  curret object;
        header is a list of header titles
        '''
        DEFAULT_FN = 'Default_CSV.xls'
        self.fn = fn
        if self.fn == None:
            self.fn = DEFAULT_FN
        self.header = headers
        self.read_row_dict =  {}
        self.write_row_dict = {}
        self.rolodex = 100

        self.restval = 'N/A'
        try:
            with open(self.fn, 'w', newline='') as sheet:
                writer = csv.DictWriter(sheet,
                                        fieldnames=self.header,
                                        restval=self.restval)
                writer.writeheader()
        except:
            pass
        
    def read(self, fn = None):
        '''
        Take rows of lists and convert store them in self.read_row_dict as
        ID: [row]
        '''
        if fn == None:
            fn = self.fn
        with open(fn, 'r') as content:
            csv_reader = csv.reader(content)
            for row in csv_reader:
                _ID = row[0]
                if _ID == 'ID':
                    self.read_row_dict['HEADERS'] = row
                else:
                    self.read_row_dict[_ID] = row
        return self.read_row_dict



    def write(self,entries, mode='a'):
        '''
        Takes a dic of dic's for entries, rolodex# is unique ID and key for
        outer dict, inner dict is header:entry paired values for rows in sheet
        '''
        header = self.header
        with open(self.fn, mode, newline='') as csvfile:
            writer = csv.DictWriter(csvfile,
                                   fieldnames = header,
                                   restval = self.restval)
            for ID, row in entries.items():
                self.write_row_dict[ID] = row
                writer.writerow(row)

    def load(self, fn):
        self.read(fn)
        for key, value in self.read_row_dict.items():
            print(dict(value))
    
            
    def find(self, value, header):
        '''
        Isolate cell wit a certain value under a specfic header,
        return all rows with that value under a common header
        '''
        fn = self.fn
        with open(fn,'r') as file:
            reader = csv.DictReader(file)
            rows = [row for row in reader if row[header] == value]
        for row in rows:
            yield row

    def remove_row(self, ID = None, value = None):
        '''
        Remove a row or rows based on ID number or value
        ID takes a list of unique ID's
        Value is a string value in any given cell
        If ID and value are provided, only
        '''
        
        if value:
            keys = []
            for key, row in self.row_dict.items():
                if ID == row['ID'] or value in row.values():
                    keys.append(key)
            for key in keys:
                self.row_dict.pop(key)

        else:
            if type(ID) is list:
                for key in ID:
                    if key in self.row_dict.keys():
                        self.row_dict.pop(key)
            else:
                if type(ID) is str:
                    if ID in self.row_dict.keys():
                        self.row_dict.pop(ID)
                    

        with open(self.fn, 'w', newline='') as sheet:
            writer = csv.DictWriter(sheet,
                                    fieldnames=self.header,
                                    restval=self.restval)
            writer.writeheader()
        self.write(self.row_dict, self.header)

        return self
            
    def remove_col(self, header):
            
        '''
        Remove data in all rows under a given header name, and the header itself
        '''
        self.header.remove(header)
        for key, value in self.row_dict.items():
            if header in value.keys():
                value.pop(header)
        with open(self.fn, 'w', newline='') as sheet:
            writer = csv.DictWriter(sheet,
                                    fieldnames=self.header,
                                    restval=self.restval)
            writer.writeheader()
        self.write(self.row_dict, self.header)
        return self

    def remove_cell(self, ID, header):
        '''
        Delete a value from a sinlge header in a given row, or list of rows
        '''
        for key, row in self.row_dict.items():
            if key in ID:
                self.row_dict[key].pop(header)

        with open(self.fn, 'w', newline='') as sheet:
            writer = csv.DictWriter(sheet,
                                    fieldnames=self.header,
                                    restval=self.restval)
            writer.writeheader()
        self.write(self.row_dict, self.header)
        content = self.read()
        
        return self
       
      

if __name__ == '__main__':
  '''
  A simple set of examples to demonstrate class functions,
  adjust comments below to activate various function calls,
  modify values below to adjust outcomes
  '''

  func = None
  #func = 'remove_row_ID'
  #func = 'remove_row_mult_ID'
  #func = 'remove_row_ID_value'
  #func = 'remove_row_value'
  #func = 'remove_col'
  #func = 'remove_cell'
  #func = 'find'

  
    
  
  FILE = 'My_demonstration_file.xls'
  

  HEADERS = ['ID', 'NAME', 'AGE', 'BREED', 'COLOR']
  VALUES = {'123':
            {'ID':'123',
            'NAME': 'Henry',
            "AGE": '37',
            "BREED":'keiglmeister',
            'COLOR':'brown'},
            '234':
            {'ID':'234',
            'NAME': 'Bobby',
            "AGE": '78',
            "BREED":'Lampoonian',
            'COLOR':'red'},
            '345':
            {'ID':'345',
            'NAME': 'Larry',
            "AGE": '12',
            "BREED":'Fartooner',
            'COLOR':'red'},
            '456':
            {'ID':'456',
            'NAME': 'Shane',
            "AGE": '66',
            "BREED":'Pythonist',
            'COLOR':'blue'}
            }
  
  sheet = Spreadsheet(HEADERS, FILE)
  info = os.stat(sheet.fn)
  print('pre',info.st_size)
  
      
  sheet.write(VALUES, mode = 'a')
  info = os.stat(sheet.fn)
  print('post', info.st_size)
 

  
  def print_pretty(func = None):
      '''
      A pleasantly formatted way to see simple examples of class functions
      results without opening a spreadsheet app, if demo file is opened in
      spread sheet app, results should match visual out come from here
      '''

      sheet = Spreadsheet(HEADERS)

      sheet.write(VALUES, HEADERS)
      print('wroten to:', info.st_size)
      content = sheet.read()

      if func == None:
          content = sheet.read()
          print('\n\nUnmodified Spreadsheet Readout')


      elif func == 'remove_row_ID':
          sheet.remove_row(ID = '234')
          content = sheet.read()
          print("\n\nResults of: sheet.remove_row(ID = '234')")

      elif func == 'remove_row_mult_ID':
          sheet.remove_row(ID = ['123', '345'])
          content = sheet.read()
          print("\n\nResults of: sheet.remove_row(ID = ['123', 345])")

      elif func == 'remove_row_ID_value':
          sheet.remove_row(ID = '456', value = 'red')
          content = sheet.read()
          print("\n\nResults of: sheet.remove_row(ID='456', value='red')")

      elif func == 'remove_row_value':
          sheet.remove_row(value = 'red')
          content = sheet.read()
          print("\n\nResults of: sheet.remove_row(value = 'red')")

      elif func == 'remove_col':
          sheet.remove_col('AGE')
          content = sheet.read()
          print('\\nResults of: sheet.remove_col("AGE")')
          for row in content:
              values = {key:value for key, value in zip(HEADERS, row)}
              print('-'*48)
              print('|{ID:^5} | {NAME:^6} | {BREED:^12} | {COLOR:^6}|'.
                    format(**values))
              print('-'*48)     

      elif func == 'remove_cell':
          sheet.remove_cell('123', 'BREED')
          content = sheet.read()
          print("\n\nResults of: sheet.remove_cell('123', 'BREED')")

      elif func == 'find':
          sheet.find('red', 'COLOR')
          content = sheet.read()
          print("Results of: sheet.find('red', 'COLOR')")
          
    

      print('-'*48+'\n'+'-'*48)
      for row in content:
          values =  {key:value for key, value in zip(HEADERS, row)}
          
          print('|{ID:^5} | {NAME:^6} | {AGE:^5} | {BREED:^12} | {COLOR:^6}|'.
                format(**values))
          print('-'*48)
      


  #print_pretty(func = func)

