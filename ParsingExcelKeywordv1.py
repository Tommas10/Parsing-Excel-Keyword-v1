#!/usr/bin/env python

#Small automation Python script- Parsing Excel keyword.
#Created by Tommas Huang 
#Created date: 2019-06-16

import xlrd
#xlrd is a module that allows Python to read data from Excel files.
from operator import itemgetter
#The operator module exports a set of efficient functions corresponding to the intrinsic operators of Python.
#operator.itemgetter(n) constructs a callable that assumes an iterable object (e.g. list, tuple, set) as input, and fetches the n-th element out of it.

keyword = "AMPICLOX"
#Search Excel keyword.

workbook = xlrd.open_workbook(r"/Users/tommashuang/Documents/test1.xls")
#Open search Excel path.
sheet = workbook.sheet_by_index(0)
rows = [sheet.row(row) for row in range(sheet.nrows)]   
# Read in all rows

rows = [row for row in rows if keyword in ' '.join(str(col.value) for col in itemgetter(2, 3)(row))]    
# Filter rows containing keyword somewhere

for row in rows:
    print(row)
