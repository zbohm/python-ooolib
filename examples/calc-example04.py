#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create the document
doc = ooolib.Calc()

# Set Column Width
doc.set_column_property(1, 'width', '0.5in')
doc.set_column_property(2, 'width', '1.0in')
doc.set_column_property(3, 'width', '1.5in')

# Set Row Height
doc.set_row_property(1, 'height', '0.5in')
doc.set_row_property(2, 'height', '1.0in')
doc.set_row_property(3, 'height', '1.5in')

# Fill in Cell Data
doc.set_cell_value(1, 1, "string", "0.5in x 0.5in")
doc.set_cell_value(2, 2, "string", "1.0in x 1.0in")
doc.set_cell_value(3, 3, "string", "1.5in x 1.5in")

# Write out the document
doc.save("calc-example04.ods")
