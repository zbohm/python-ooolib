#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create your document
doc = ooolib.Calc()

# Set values.  The values are set in column, row order, but the values are
# not in the traditional "A5" style format.  Instead we require two integers.
# set_cell_value(col, row, datatype, value)
doc.set_cell_value(1, 1, "string", 'Alignment')

# Set Bold
doc.set_cell_property('bold', True)

doc.set_row_property(3, 'height', '0.5in')
doc.set_row_property(4, 'height', '0.5in')
doc.set_row_property(5, 'height', '0.5in')

# valign Labels
doc.set_cell_value(1, 3, "string", 'top')
doc.set_cell_value(1, 4, "string", 'middle')
doc.set_cell_value(1, 5, "string", 'bottom')

# halign Labels
doc.set_cell_value(2, 2, "string", 'left')
doc.set_cell_value(3, 2, "string", 'center')
doc.set_cell_value(4, 2, "string", 'right')
doc.set_cell_value(5, 2, "string", 'justify')
doc.set_cell_value(6, 2, "string", 'filled')

# Unset Bold
doc.set_cell_property('bold', False)

# Fill in aligned properties
# Vertical Align top
doc.set_cell_property('valign', 'top')

doc.set_cell_property('halign', 'left')
doc.set_cell_value(2, 3, "string", 'x')
doc.set_cell_property('halign', 'center')
doc.set_cell_value(3, 3, "string", 'x')
doc.set_cell_property('halign', 'right')
doc.set_cell_value(4, 3, "string", 'x')
doc.set_cell_property('halign', 'justify')
doc.set_cell_value(5, 3, "string", 'x')
doc.set_cell_property('halign', 'filled')
doc.set_cell_value(6, 3, "string", 'x')

# Vertical Align middle
doc.set_cell_property('valign', 'middle')

doc.set_cell_property('halign', 'left')
doc.set_cell_value(2, 4, "string", 'x')
doc.set_cell_property('halign', 'center')
doc.set_cell_value(3, 4, "string", 'x')
doc.set_cell_property('halign', 'right')
doc.set_cell_value(4, 4, "string", 'x')
doc.set_cell_property('halign', 'justify')
doc.set_cell_value(5, 4, "string", 'x')
doc.set_cell_property('halign', 'filled')
doc.set_cell_value(6, 4, "string", 'x')

# Vertical Align bottom
doc.set_cell_property('valign', 'bottom')

doc.set_cell_property('halign', 'left')
doc.set_cell_value(2, 5, "string", 'x')
doc.set_cell_property('halign', 'center')
doc.set_cell_value(3, 5, "string", 'x')
doc.set_cell_property('halign', 'right')
doc.set_cell_value(4, 5, "string", 'x')
doc.set_cell_property('halign', 'justify')
doc.set_cell_value(5, 5, "string", 'x')
doc.set_cell_property('halign', 'filled')
doc.set_cell_value(6, 5, "string", 'x')

# Set Default Alignments
doc.set_cell_property('valign', 'default')
doc.set_cell_property('halign', 'default')

# Save the document to the file you want to create
doc.save("calc-example08.ods")
