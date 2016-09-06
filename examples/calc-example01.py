#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create your document
doc = ooolib.Calc()

# Set values.  The values are set in column, row order, but the values are
# not in the traditional "A5" style format.  Instead we require two integers.
# set_cell_value(col, row, datatype, value)
for row in range(1, 9):
	for col in range(1, 9):
		doc.set_cell_value(col, row, "float", col * row)

# Save the document to the file you want to create
doc.save("calc-example01.ods")
