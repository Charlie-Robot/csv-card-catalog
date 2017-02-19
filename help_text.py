display = {
    
'Find':
'''
Find:

Takes a single Header value, and a single Cell value. Search currently open \
file for data rows that match User defined criterion. Card que is replaced \
with only the matching data rows.  If no matching rows are found, que is \
unaltered.  Save as new file, or overwrite an existing file.\
''',
'Edit Cell':
'''
Edit Cell:

Allows you to alter the value of a single cell, under a given Header for one \
or more rows at a time. Rolodex and Header are Required fields, if New Cell Value \
is left blank, specified cells will become blank. Save as new file, or overwrite \
an existing file. 

Rolodex:
Takes a single Rolodex ID, a range of Rolodex ID's, or an arbitrary \
combination of comma seperated single ID's and ranges.

EXAMPLES:
Single rows:
100 -OR- 100,103,110,225

Range of Rows:
150-275 -OR- 100-150, 155-200

Combination of single rows and a range of rows:
100, 115, 150-275, 300, 305-325, 290-295


Header:
Takes a single Fieldname value only

New Cell Value:
Takes a single Cell value only
''',

'Remove Row':
'''
Remove Row:

Remove an entire row and all of its contents based on a combination of Rolodex ID, \
and/or Header = Cell values. If no Rolodex is provided, a Header and Value are required;\
but if Rolodex is specified, Header and Value can be omitted. All Rows matching the \
given criteria will be deleted.

Rolodex:
Takes a single Rolodex ID, a range or Rolodex ID's, or an arbitrary \
combination of comma seperated single ID's and ranges.

EXAMPLES:
Single rows:
100 -OR- 100,103,110,225

Range of Rows:
150-275 -OR- 100-150, 155-200

Combination of single rows and a range of rows:
100, 115, 150-275, 300, 305-325, 290-295


Header:
Takes a single Fieldname value only

Cell Value:
Takes a single Cell value only
''',

'Remove Column':
'''
Remove Column:

Remove An entire column of data under a given Header name. All rows in currently open \
file will have their data removed from that cell.

Header:
Takes a single Fieldname Value only
''',

'Merge':
'''
Merge:
Combine a second existing file with currently open file. Combined files are automatically \
saved as a third file named as currently open file name + "_merged.csv".
Original two files are unaltered.

In the event that the second file has a Rolodex ID already in currently opened file, you \
will be prompted to either change the second files Rolodex ID and include the data row; \
overwrite the data row in currently open file with the row from the second file; OR exclude \
the second files conflicting row.
''',

'Read Out':
'''
Read Out:
A way to view of the currently open file content without a third party program.  Content cannot \
be directly edited or altered.  Read Out window does not reflect changes made to file while \
it is open.  
'''
}

