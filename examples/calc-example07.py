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
		doc.set_cell_value(col, row, "string", str(col * row))
		doc.set_cell_value(col, row, "annotation", "Annotation:" + str(col*row))
		doc.set_cell_value(col, row, "link", ('file:///c:\\tmp\\content.xml', 'My Link'))
doc.set_cell_value(1, 9, "link", ('http://ooolib.sourceforge.net/', 'ooolib'))

# Save the document to the file you want to create
doc.save("calc-example07.ods")
