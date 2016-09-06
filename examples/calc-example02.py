#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create the document
doc = ooolib.Calc("First")

# By default you start on Sheet1.  This has an index of 0 in ooolib.
doc.set_cell_value(1, 1, "string", "We start on \"Sheet1\"")

# Create a new sheet by passing the title.  You will automatically
# move to that sheet.
doc.new_sheet("Second")
doc.set_cell_value(1, 1, "string", "I'm on Sheet2")

# Create another one
doc.new_sheet("Sheet3")
doc.set_cell_value(1, 1, "string", "This is Sheet3")

# Move back to the first sheet
doc.set_sheet_index(0)
doc.set_cell_value(1, 2, "string", "Sheet1's index is 0")

# Move back to the second sheet
doc.set_sheet_index(1)
doc.set_cell_value(1, 2, "string", "Sheet2's index is 1")

# Write out the document
doc.save("calc-example02.ods")
