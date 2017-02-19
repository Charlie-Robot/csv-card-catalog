import unittest, os, time
from sheet import Spreadsheet

class Unitest(unittest.TestCase):

    def setUp(self):
        self.FILE = 'test.csv'
        self.FILE_1 = 'test_2.csv'
        self.FILE_2 = 'test_3.csv'
        
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
        self.FULL_VALS  = {'500': {'First' : 'Madison',
                                   'Middle' : 'Katy',
                                   'Last' : 'Wyss',
                                   'Street' :  'Marina Blvd',
                                   'Apt#' : '305',
                                   'City' : 'San Leandro',
                                   'Zip' : '90210',
                                   'Email' : 'madison@baby.com',
                                   'Occupation' : 'baby',
                                   'Employer' : 'Mom',
                                   'Filed By' : 'Igor',
                                   'Rolodex' : '500',
                                   'Member Signed' : 'yes',
                                   'Cope Signed': 'yes',
                                   'Date Signed' : '01/01/2001',
                                   'Contribution' : '10'}}
        self.MISSING_VALS  = {'100': {'First' : 'Hobart',
                                      'Middle' : 'The',
                                      'Last' : 'Great',
                                      'Street' :  'Awesome Ln',
                                      'Apt#' : '',
                                      'City' : 'Candy Land',
                                      'Zip' : '12345',
                                      'Email' : '',
                                      'Occupation' : 'Magician',
                                      'Employer' : '',
                                      'Filed By' : 'Igor',
                                      'Rolodex' : '100',
                                      'Member Signed' : 'yes',
                                      'Cope Signed': 'yes',
                                      'Date Signed' : '02/11/2012',
                                      'Contribution' : '0'}}
        self.MORE_VALS  = {'501': {'First' : 'Sheman',
                                   'Middle' : 'Tai',
                                   'Last' : 'Wyss',
                                   'Street' :  'Mallard Loop',
                                   'Apt#' : '133',
                                   'City' : 'Whitefish',
                                   'Zip' : '59937',
                                   'Email' : 'stumpy@lowdog.com',
                                   'Occupation' : 'Guard Dog',
                                   'Employer' : 'Mom',
                                   'Filed By' : 'Igor',
                                   'Rolodex' : '501',
                                   'Member Signed' : 'yes',
                                   'Cope Signed': 'yes',
                                   'Date Signed' : '12/10/2011',
                                   'Contribution' : '20'}}
        self.OVERLAP_VALS  = {'500': {'First' : 'Franklin',
                                       'Middle' : 'The',
                                       'Last' : 'Lesser',
                                       'Street' :  'Adequate Ln',
                                       'Apt#' : '',
                                       'City' : 'Kalamazoo',
                                       'Zip' : '24681',
                                       'Email' : 'not_so_great@humble.org',
                                       'Occupation' : 'Ponsy',
                                       'Employer' : 'Mom',
                                       'Filed By' : 'Igor',
                                       'Rolodex' : '500',
                                       'Member Signed' : 'yes',
                                       'Cope Signed': 'yes',
                                       'Date Signed' : '12/2/2014',
                                       'Contribution' : '0'},
                              '100': {'First' : 'Tech',
                                       'Middle' : 'The',
                                       'Last' : 'Tank',
                                       'Street' :  'Podunk Blvd',
                                       'Apt#' : '',
                                       'City' : 'Haywire Gulch',
                                       'Zip' : '11223',
                                       'Email' : 'tractor_time@hotmail.com',
                                       'Occupation' : 'Tractor',
                                       'Employer' : 'Old McDonald',
                                       'Filed By' : 'Igor',
                                       'Rolodex' : '100',
                                       'Member Signed' : 'yes',
                                       'Cope Signed': 'yes',
                                       'Date Signed' : '10/12/2004',
                                       'Contribution' : '10'}}

        self.actual_dic = {}
        self.sheet_1 = Spreadsheet(self.HEADERS,self.FILE)
        self.sheet_2 = Spreadsheet(self.HEADERS,self.FILE_1)
        self.sheet_3 = Spreadsheet(self.HEADERS,self.FILE_2)
        self.sheet_4 = Spreadsheet(self.HEADERS)
        FN, ext = os.path.splitext(self.sheet_1.FN)
        self.MERGED_FN = '{}_merged{}'.format(FN, ext)

   
    def test_basic_class_properties(self):
        #Affirm seperate objects
        self.assertNotEqual(self.sheet_1, self.sheet_2)
        self.assertNotEqual(self.sheet_1, self.sheet_3)
        self.assertNotEqual(self.sheet_2, self.sheet_3)
        #Test object sheet_1
        self.sheet_1.row_dict = {}
        self.assertEqual(self.sheet_1.FN, self.FILE)
        self.assertEqual(self.sheet_1.HEADER, self.HEADERS)
        #Test object sheet_2
        self.sheet_2.row_dict = {}
        self.assertEqual(self.sheet_2.FN, self.FILE_1)
        self.assertEqual(self.sheet_2.HEADER, self.HEADERS)
        #Test object sheet_3
        self.sheet_3.row_dict = {}
        self.assertEqual(self.sheet_3.FN, self.FILE_2)
        self.assertEqual(self.sheet_3.HEADER, self.HEADERS)
        #Test default file naming convention
        #Spreadsheet created with unspecified file name is auto name with current date 01.01.2016
        self.sheet_4.write(self.FULL_VALS)
        DATE_FORMAT = '%m/%d/%Y'
        TODAY = time.strftime(DATE_FORMAT)
        EXPECTED_DEFAULT_FN = TODAY.replace('/', '.') + '.csv'
        self.assertTrue(os.path.isfile(EXPECTED_DEFAULT_FN))
        self.assertEqual(EXPECTED_DEFAULT_FN, self.sheet_4.FN)
        #If sheet with current date as file name already exists,
        #append (n) to default file name, first duplicate is (1), every duplicate file name
        #that follows is n+1 ==> 01.01.2010.csv => 01.01.2010(1)/csv => 01.01.2010(2).csv
        ACTUAL_DUPLICATES = []
        for n in range(1,21):     
            EXPECTED_DUPLICATE_FN = TODAY.replace('/','.') + '({})'.format(n) + '.csv'
            #create 20 duplicate file names
            duplicate_sheet_name = Spreadsheet(self.HEADERS)
            duplicate_sheet_name.write(self.FULL_VALS)
            ACTUAL_DUPLICATES.append(duplicate_sheet_name.FN)
            self.assertTrue(os.path.isfile(EXPECTED_DUPLICATE_FN))
            self.assertNotEqual(EXPECTED_DEFAULT_FN, duplicate_sheet_name.FN)
            self.assertEqual(EXPECTED_DUPLICATE_FN, duplicate_sheet_name.FN)
        for FN in ACTUAL_DUPLICATES:
            os.remove(FN)
        #Check csv RESTVAL
        self.assertEqual('-', self.sheet_1.RESTVAL)
           
            
            
        
    
    def test_write(self):
        '''
        Test a single row of data
        '''
        self.assertTrue(len(self.sheet_1.READ_DICT) == 0)
        self.assertTrue(len(self.sheet_1.WRITE_DICT) == 0)
        
        
        
        #Write a single line of data below Headers
        self.sheet_1.write(self.FULL_VALS)
        #Number of Values does not exceed number of headers
        self.assertTrue(len(self.FULL_VALS.values()) <= len(self.HEADERS))
        #write() created a valid .csv file
        FN, ext = os.path.splitext(self.sheet_1.FN)
        self.assertTrue(os.path.isfile(self.FILE), '{} is not a file'.format(self.FILE))
        self.assertTrue(ext == '.csv')
        
        with open(self.FILE, 'r') as content:
           for row in content:
               row = row.split(',')
               self.assertNotEqual(row, self.HEADERS)
               row_as_dict = {header:value for header, value in zip(self.HEADERS, row)}
               
        for key, row in self.FULL_VALS.items():
            self.assertEqual(sorted(row), sorted(row_as_dict))
            self.assertEqual(row_as_dict['Rolodex'], key)
        #write() should remove key, value pairs from entries if value == ''
        #these removed values will be replaced by csv.DictWriter's restval
        missing_vals = ['Apt#', 'Employer', 'Email']
        for header in missing_vals:
            self.assertTrue(header in self.MISSING_VALS['100'].keys())
        self.sheet_2.write(self.MISSING_VALS)
        for header in missing_vals:
            self.assertTrue(header not in self.MISSING_VALS['100'].keys())
                                                           
        
    def test_read(self):
        '''
        Read in rows of data into a dict of lists, update Spreadsheet Rolodex
        to highest Rolodex.
        '''
        EXPECTED = {'501': ['Sheman',
                            'Tai',
                            'Wyss',
                            'Mallard Loop',
                            '133',
                            'Whitefish',
                            '59937',
                            'stumpy@lowdog.com',
                            'Guard Dog',
                            'Mom',
                            'Igor',
                            '501',
                            'yes',
                            'yes',
                            '12/10/2011',
                            '20'],
                    '500' :['Madison',
                            'Katy',
                            'Wyss',
                            'Marina Blvd',
                            '305',
                            'San Leandro',
                            '90210',
                            'madison@baby.com',
                            'baby',
                            'Mom',
                            'Igor',
                            '500',
                            'yes',
                            'yes',
                            '01/01/2001',
                            '10']}
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MORE_VALS)
        test_sheet = Spreadsheet(self.HEADERS, 'test_sheet.csv')
        self.assertTrue(os.path.isfile(self.sheet_1.FN))
        default_rolodex = 100
        self.assertEqual(default_rolodex, test_sheet.ROLODEX)
        test_sheet.read(self.sheet_1.FN)
        self.assertEqual(test_sheet.ROLODEX, 501)
        
        for key, row in EXPECTED.items():
            self.assertTrue(key in test_sheet.READ_DICT.keys())
            self.assertEqual(test_sheet.READ_DICT[key], row)
        self.assertEqual(EXPECTED, test_sheet.READ_DICT)

      
    

    def test_empty_cell(self):
        '''
        Undefined headers values  in input dic should be filled with a "N/A"
        if corresponding spread sheet cell
        '''
        
        EXPECTED  = ['Hobart',
                     'The',
                     'Great',
                     'Awesome Ln',
                     '-',
                     'Candy Land',
                     '12345',
                     '-',
                     'Magician',
                     '-',
                     'Igor',
                     '100',
                     'yes',
                     'yes',
                     '02/11/2012',
                     '0']

        #Write a single line of data below Headers
        self.sheet_1.write(self.MISSING_VALS)

        #Read Spreadsheet row
        self.sheet_1.read()
        for key, row in self.sheet_1.READ_DICT.items():
            for indx, value in enumerate(row):
                self.assertEqual(EXPECTED[indx], value,"Index: {} not 'N/A'".format(indx))
       

    def test_find(self):
        '''
        spreadsheet.find(header, value), return a list of rows
        where header=value in current working file
        '''
        #write two rows into spreadsheet
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        #simulate a loaded spreadsheet
        self.sheet_1.read()
        self.assertEqual(2, len(self.sheet_1.WRITE_DICT))
        self.assertTrue("Last" in self.HEADERS)

        #No matching value does not alter sheet.WRITE_DICT
        self.sheet_1.find('Last', 'Smith')
        self.assertEqual(2, len(self.sheet_1.WRITE_DICT))

        #Both rows have matching value
        self.assertTrue('Filed By' in self.HEADERS)
        self.sheet_1.find('Filed By', 'Igor')
        self.assertEqual(2, len(self.sheet_1.WRITE_DICT))
        self.assertTrue('100' in self.sheet_1.WRITE_DICT.keys())
        self.assertTrue('500' in self.sheet_1.WRITE_DICT.keys())

        #Single row with matching value
        self.sheet_1.find('Last', 'Wyss')
        self.assertEqual(1, len(self.sheet_1.WRITE_DICT))
        self.assertTrue(self.FULL_VALS == self.sheet_1.WRITE_DICT)
    
                
    def test_remove_single_row_by_rolodex_only(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove a single row of data according to a Rolodex number
        self.sheet_1.remove_row('501','','')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' in ALTERED_KEYS)
        self.assertTrue('500' in ALTERED_KEYS)
        self.assertTrue('501' not in ALTERED_KEYS)
        
    def test_remove_multiple_rows_by_rolodex_only(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.sheet_2.write(self.FULL_VALS)
        self.sheet_2.write(self.MISSING_VALS)
        self.sheet_2.write(self.MORE_VALS)
        self.sheet_2.read()
        self.sheet_3.write(self.FULL_VALS)
        self.sheet_3.write(self.MISSING_VALS)
        self.sheet_3.write(self.MORE_VALS)
        self.sheet_3.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove all rows in a given range of Rolodex numbers
        self.sheet_1.remove_row('101-600','','')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)
        self.assertTrue('501' not in ALTERED_KEYS)

        #Remove all rows in a range of Rolodex numbers combined with indivudual
        #Rolodex numbers outside of range
        ORIGINAL_KEYS = self.sheet_2.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.sheet_2.remove_row('0-100, 500', '', '')
        ALTERED_KEYS = self.sheet_2.WRITE_DICT.keys()
        self.assertTrue('100' not in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)

        #Remove all rows in a comma seperated set of rolodex numbers
        ORIGINAL_KEYS = self.sheet_3.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.sheet_3.remove_row('100, 500', '', '')
        ALTERED_KEYS = self.sheet_3.WRITE_DICT.keys()
        self.assertTrue('100' not in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)
        
        
    def test_remove_single_row_by_rolodex_plus_header_value(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove all rows in a given range of Rolodex numbers where header == value
        self.sheet_1.remove_row('000-600','Last','Great')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' not in ALTERED_KEYS)
        self.assertTrue('500' in ALTERED_KEYS)
        self.assertTrue('501' in ALTERED_KEYS)

    def test_remove_multiple_rows_by_rolodex_plus_header_value(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove all rows in a given range of Rolodex numbers
        self.sheet_1.remove_row('000-600','Last','Wyss')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)
        self.assertTrue('501' not in ALTERED_KEYS)

    def test_remove_multiple_rows_by_header_value_only(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove all rows in a given range of Rolodex numbers
        self.sheet_1.remove_row('','Last','Wyss')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)
        self.assertTrue('501' not in ALTERED_KEYS)
        
    def test_remove_single_row_by_header_value_only(self):
        #Create Spreadsheet with3 rows of data
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        self.assertEqual(3, len(self.sheet_1.WRITE_DICT))
        ORIGINAL_KEYS = self.sheet_1.WRITE_DICT.keys()
        self.assertTrue('100' in ORIGINAL_KEYS)
        self.assertTrue('500' in ORIGINAL_KEYS)
        self.assertTrue('501' in ORIGINAL_KEYS)

        #Remove all rows in a given range of Rolodex numbers
        self.sheet_1.remove_row('','Contribution','10')
        ALTERED_KEYS = self.sheet_1.WRITE_DICT.keys()

        self.assertEqual(ORIGINAL_KEYS, ALTERED_KEYS)
        self.assertTrue('100' in ALTERED_KEYS)
        self.assertTrue('500' not in ALTERED_KEYS)
        self.assertTrue('501' in ALTERED_KEYS)

    def test_edit_cell(self):
        '''
        edit_cell(rolodex, header, new_value, [over_ride = False])
        Alter the value of a single cell in a row or rows of Data,
        spreadsheet.edit_cell('100', 'Header', 'New_Value'),
        -or-
        spreadsheet.edit_cell('100,101,103', 'Header', 'New_Value'),
        -or-
        spreadsheet.edit_cell('100, 200-600', 'Header', 'New_Value')
        Rolodex unalterable unless over_ride == True
        '''
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        #Unaltered data row
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Contribution'],
                         self.MISSING_VALS['100']['Contribution'])
        #Call spreadsheet.edit_cell() on single row
        self.sheet_1.edit_cell('100', 'Contribution', '30')
        #Spreadsheet altered value of given cell
        self.assertNotEqual(self.sheet_1.WRITE_DICT['100']['Contribution'],
                            self.MISSING_VALS['100']['Contribution'])
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Contribution'], '30')
        
        #Call spreadsheet.edit_cell() on multiple rows
        #Test unaltered values 
        self.assertEqual(self.MISSING_VALS['100']['Filed By'],
                         self.sheet_1.WRITE_DICT['100']['Filed By'])
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Filed By'],'Igor')
        self.assertEqual(self.FULL_VALS['500']['Filed By'],
                         self.sheet_1.WRITE_DICT['500']['Filed By'])
        self.assertEqual(self.sheet_1.WRITE_DICT['500']['Filed By'],'Igor')
        self.assertEqual(self.MORE_VALS['501']['Filed By'],
                         self.sheet_1.WRITE_DICT['501']['Filed By'])
        self.assertEqual(self.sheet_1.WRITE_DICT['501']['Filed By'],'Igor')
        #Alter single cell in multiple Data rows
        #individual rolodex ID's
        self.sheet_1.edit_cell('100,500,501', 'Filed By', 'Jack')
        #All 3 rows in spreadsheet altered 'Filed By' value
        #Test altered values
        self.assertNotEqual(self.MISSING_VALS['100']['Filed By'],
                         self.sheet_1.WRITE_DICT['100']['Filed By'])
        self.assertEqual(self.MISSING_VALS['100']['Filed By'], 'Igor')
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Filed By'],'Jack')
        
        self.assertNotEqual(self.FULL_VALS['500']['Filed By'],
                         self.sheet_1.WRITE_DICT['500']['Filed By'])
        self.assertEqual(self.FULL_VALS['500']['Filed By'],'Igor')
        self.assertEqual(self.sheet_1.WRITE_DICT['500']['Filed By'],'Jack')
        
        self.assertNotEqual(self.MORE_VALS['501']['Filed By'],
                         self.sheet_1.WRITE_DICT['501']['Filed By'])
        self.assertEqual(self.MORE_VALS['501']['Filed By'],'Igor')
        self.assertEqual(self.sheet_1.WRITE_DICT['501']['Filed By'],'Jack')
        #range of rolodex ID's
        self.sheet_1.edit_cell('0-100, 501', 'Employer', 'Dr. Frankenstein')
        #test altered range
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Employer'],'Dr. Frankenstein')
        self.assertNotEqual(self.MORE_VALS['501']['Employer'],
                         self.sheet_1.WRITE_DICT['501']['Employer'])
        self.assertEqual(self.MORE_VALS['501']['Employer'],'Mom')
        self.assertEqual(self.sheet_1.WRITE_DICT['501']['Employer'],'Dr. Frankenstein')
        

        #Rolodex can only be changed if over_ride  == True
        self.assertTrue('100' in self.sheet_1.WRITE_DICT.keys())
        #edit_cell called without over_ride invoked, rolodex and WRITE_DICT[INDEX] should not change
        self.sheet_1.edit_cell('100', 'Rolodex', '222')
        self.assertTrue('100' in self.sheet_1.WRITE_DICT.keys())
        self.assertEqual(self.sheet_1.WRITE_DICT['100']['Rolodex'], '100')
        #edit_cell called with over_ride invoked, rolodex and WRITE_DICT[INDEX] should change
        self.sheet_1.edit_cell('100', 'Rolodex', '222', over_ride = True)
        self.assertTrue('100' not in self.sheet_1.WRITE_DICT.keys())
        self.assertEqual(self.sheet_1.WRITE_DICT['222']['Rolodex'], '222')
                
        
    def test_remove_column(self):
        '''
        All values under under a given header are removed without removing the header itself
        '''
        HEADER = "Apt#"
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_1.write(self.MORE_VALS)
        self.sheet_1.read()
        for indx, row in self.sheet_1.WRITE_DICT.items():
            self.assertTrue(HEADER in row.keys())
        #call spreadsheet.remove_col()
        self.sheet_1.remove_col(HEADER)
        
        for indx, row in self.sheet_1.WRITE_DICT.items():
            self.assertTrue(HEADER in row.keys())
            self.assertTrue(row[HEADER] == '')
   
    def test_basic_merge(self):
        '''
        Combine current working spreadsheet with second existing sheet, results written to third
        .csv file, original files are unaltered
        '''
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_2.write(self.OVERLAP_VALS)
        self.sheet_3.write(self.MORE_VALS)
        self.sheet_1.read()
        self.sheet_2.read()
        self.sheet_3.read()
        self.assertTrue(os.path.isfile(self.sheet_1.FN))
        self.assertTrue(os.path.isfile(self.sheet_2.FN))
        self.assertTrue(os.path.isfile(self.sheet_3.FN))
        self.assertNotEqual(self.sheet_1.FN, self.sheet_2.FN)
        self.assertNotEqual(self.sheet_1.FN, self.sheet_3.FN)
        self.assertNotEqual(self.sheet_2.FN, self.sheet_3.FN)
        sheet_1_keys = self.sheet_1.WRITE_DICT.keys()
        sheet_2_keys = self.sheet_2.WRITE_DICT.keys()
        sheet_3_keys = self.sheet_3.WRITE_DICT.keys()
        
        #Merge two files without overlapping Roldex Numbers
        FN, ext = os.path.splitext(self.sheet_1.FN)
        for rolodex in self.sheet_1.WRITE_DICT.keys():
            self.assertTrue(rolodex not in self.sheet_3.WRITE_DICT.keys())
        self.assertFalse(os.path.isfile(self.MERGED_FN))
        
        #the merge
        self.sheet_1.merge(self.sheet_3.FN)
        #merged file was created
        self.assertTrue(os.path.isfile(self.MERGED_FN))
        self.sheet_4.read(self.MERGED_FN)
        #merged file contains rows of both sheet_1 and sheet_3
        for rolodex, row in self.sheet_1.WRITE_DICT.items():
            self.assertTrue(rolodex in self.sheet_4.WRITE_DICT.keys())
            self.assertEqual(self.sheet_1.WRITE_DICT[rolodex],
                             self.sheet_4.WRITE_DICT[rolodex])
        for rolodex, row in self.sheet_3.WRITE_DICT.items():
            self.assertTrue(rolodex in self.sheet_4.WRITE_DICT.keys())
            self.assertEqual(self.sheet_3.WRITE_DICT[rolodex],
                             self.sheet_4.WRITE_DICT[rolodex])
        #Original files are unaltered
        for rolodex in self.sheet_1.WRITE_DICT.keys():
            self.assertTrue(rolodex not in self.sheet_3.WRITE_DICT.keys())
        self.assertEqual(self.sheet_1.WRITE_DICT.keys(), sheet_1_keys)
        for rolodex in self.sheet_3.WRITE_DICT.keys():
            self.assertTrue(rolodex not in self.sheet_1.WRITE_DICT.keys())
        self.assertEqual(self.sheet_3.WRITE_DICT.keys(), sheet_3_keys)
       
         

    def test_merge_overlapping_ID_change_all_duplicate_IDs(self):
        '''
        Test onverlapping rolodex ID's
        If duplicate ID's found, GUI window prompts to change conflicting ID
        to 1 higher than current known ID (or not), with option to do same for
        all remaining conflicting ID's,IF NOT, Second GUI prompts to overwrite
        secondary ID and corrosponding row in new combined file (original file unaltered)
        with option to overwrite all remaining ID's, GUI's suppressed for testing
        '''
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_2.write(self.OVERLAP_VALS)
        self.sheet_1.read()
        self.sheet_2.read()
        self.assertTrue(os.path.isfile(self.sheet_1.FN))
        self.assertTrue(os.path.isfile(self.sheet_2.FN))
        self.assertNotEqual(self.sheet_1.FN, self.sheet_2.FN)
        sheet_1_keys = self.sheet_1.WRITE_DICT.keys()
        sheet_2_keys = self.sheet_2.WRITE_DICT.keys()
        
        
        #Conditions to pass to Spreadsheet.merge()
        conditions = {'TESTING' : True,
                      'CHANGE_YES_ALL' : True,
                      'CHANGE_YES_ONE' : False,
                      'CHANGE_NO_ALL' : False,
                      'OVERWRITE_YES_ALL' : False,
                      'OVERWITE_NO_ALL' : False}
    
        #Merge sheet_1 & sheet_2, sheet_2 have duplicate Rolodex Values
        for key in self.sheet_2.WRITE_DICT.keys():
            self.assertTrue(key in self.sheet_1.WRITE_DICT.keys())
        #MAX rolodex of both sheets being pre_merge,
        MAX_ID = max(self.sheet_1.ROLODEX, self.sheet_2.ROLODEX)
        #Expected newly merged file does not exist yet
        self.assertFalse(os.path.isfile(self.MERGED_FN))
        #The Merge
        ORIGINAL_ROLODEX = self.sheet_1.ROLODEX
        msg = self.sheet_1.merge(self.sheet_2.FN, **conditions)
        #Original sheets are unaltered
        self.assertEqual(self.sheet_1.WRITE_DICT.keys(), sheet_1_keys)
        self.assertEqual(self.sheet_2.WRITE_DICT.keys(), sheet_2_keys)
        #Expected merged file was created
        self.assertTrue(os.path.isfile(self.MERGED_FN))
        
        self.sheet_1.load(self.MERGED_FN)
        #merged file should have 4 entries, Rolodex ID's from sheet_2 being += 1
        #to highest known ID
        merged = Spreadsheet(self.HEADERS, fn = self.MERGED_FN)
        merged.read()

        #merged file should have 4 entries

        self.assertEqual(4, len(merged.WRITE_DICT.keys()))
        #Test all ID's from sheet_1 are in merged file unaltered
        for ID in self.sheet_1.WRITE_DICT.keys():
            self.assertTrue(ID in merged.WRITE_DICT.keys())
        #Test all ID's from sheet_2 are Altered and added to merged
        for ID in self.sheet_2.WRITE_DICT.keys():
            MAX_ID += 1
            self.assertTrue(str(MAX_ID) in merged.WRITE_DICT.keys())
        #Rows with Altered ID's, merged.WRITE_DICT[ID]['Rolodex'] reflects changed value
        for ID in merged.WRITE_DICT.keys():
            self.assertEqual(merged.WRITE_DICT[ID]['Rolodex'], ID)
        #Currently working sheet's Rolodex ID is updated to reflect new higher Rolodex
        self.assertNotEqual(ORIGINAL_ROLODEX, self.sheet_1.ROLODEX)
        self.assertEqual(self.sheet_1.ROLODEX, merged.ROLODEX)
        
  
    def test_merge_overlapping_ID_change_single_duplicate_IDs(self):
        '''
        If duplicate ID is not changed, data row with duplicate ID is not
        added to newly formed file, Test is limited to only changing first
        duplicate ID, actual function should allow to skip first found duplicate
        but change following duplicates, GUI's suppressed for testing
        '''
        self.sheet_1.write(self.FULL_VALS)
        self.sheet_1.write(self.MISSING_VALS)
        self.sheet_2.write(self.OVERLAP_VALS)
        self.sheet_1.read()
        self.sheet_2.read()
        self.assertTrue(os.path.isfile(self.sheet_1.FN))
        self.assertTrue(os.path.isfile(self.sheet_2.FN))
        self.assertNotEqual(self.sheet_1.FN, self.sheet_2.FN)
        sheet_1_keys = self.sheet_1.WRITE_DICT.keys()
        sheet_2_keys = self.sheet_2.WRITE_DICT.keys()
        
    
        #Conditions to pass to Spreadsheet.merge()
        conditions = {'TESTING' : True,
                      'CHANGE_YES_ALL' : False,
                      'CHANGE_YES_ONE' : True,
                      'CHANGE_NO_ALL' : False,
                      'OVERWRITE_YES_ALL' : False,
                      'OVERWRITE_YES_ONE' : False,
                      'OVERWRITE_NO_ALL' : False,
                      'OVERWRITE_NO_ONE' : False}
    
        #Merge sheet_1 & sheet_2, sheet_2 have duplicate Rolodex Values
        for key in self.sheet_2.WRITE_DICT.keys():
            self.assertTrue(key in self.sheet_1.WRITE_DICT.keys())
        #MAX rolodex of both sheets pre_merge,
        MAX_ID = max(self.sheet_1.ROLODEX, self.sheet_2.ROLODEX)
        #Expected newly merged file does not exist yet
        self.assertFalse(os.path.isfile(self.MERGED_FN))
        #The Merge
        ORIGINAL_ROLODEX = self.sheet_1.ROLODEX
        self.sheet_1.merge(self.sheet_2.FN, **conditions)
        #Original sheets are unaltered
        self.assertEqual(self.sheet_1.WRITE_DICT.keys(), sheet_1_keys)
        self.assertEqual(self.sheet_2.WRITE_DICT.keys(), sheet_2_keys)
        #Expected merged file was created
        self.assertTrue(os.path.isfile(self.MERGED_FN))
        self.sheet_1.load(self.MERGED_FN)
        
        #merged file should have 3 entries, Rolodex ID from sheet_2 being += 1
        #to highest known ID
        merged = Spreadsheet(self.HEADERS, self.MERGED_FN)
        merged.read()
        self.assertEqual(max(merged.WRITE_DICT.keys()), str(MAX_ID + 1))
        #merged file should have 3 entries
        self.assertEqual(3, len(merged.WRITE_DICT.keys()))
        #Test all ID's from sheet_1 are in merged file unaltered
        for ID in self.sheet_1.WRITE_DICT.keys():
            self.assertTrue(ID in merged.WRITE_DICT.keys())
        #Test sinlge (lowest ID) from sheet_2 are Altered and added to merged
        EXPECTED_ROW = self.sheet_2.WRITE_DICT['100']
        EXPECTED_ROW['Rolodex'] = str(MAX_ID + 1)
        self.assertTrue(EXPECTED_ROW in merged.WRITE_DICT.values())
        self.assertEqual(merged.WRITE_DICT[str(MAX_ID + 1)], EXPECTED_ROW)
        self.assertEqual(merged.WRITE_DICT['501']['Rolodex'], '501')
        #Rows with Altered ID's, merged.WRITE_DICT[ID]['Rolodex'] reflects changed value
        for ID in merged.WRITE_DICT.keys():
            self.assertEqual(merged.WRITE_DICT[ID]['Rolodex'], ID)
        #Currently working sheet's Rolodex ID is updated to reflect new higher Rolodex
        self.assertNotEqual(ORIGINAL_ROLODEX, self.sheet_1.ROLODEX)
        self.assertEqual(self.sheet_1.ROLODEX, merged.ROLODEX)

 




            
    def tearDown(self):
        for FN in [self.FILE, self.FILE_1, self.FILE_2, self.sheet_4.FN, self.MERGED_FN]:
            if os.path.isfile(FN):
                os.remove(FN)
        
        
        

        
    
        

if __name__ == '__main__':
    unittest.main()
