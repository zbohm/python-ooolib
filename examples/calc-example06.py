#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create your document
doc = ooolib.Calc()

for row in range(1, 9):
	doc.set_cell_value(1, row, "float", row)

doc.set_cell_value(2, 1, 'string', 'AVERAGE')
doc.set_cell_value(2, 2, 'string', 'MIN')
doc.set_cell_value(2, 3, 'string', 'MAX')
doc.set_cell_value(2, 4, 'string', 'SUM')
doc.set_cell_value(2, 5, 'string', 'SQRT')
doc.set_cell_value(2, 6, 'string', 'Condition')

doc.set_cell_value(3, 1, 'formula', '=AVERAGE(A1:A8)')
doc.set_cell_value(3, 2, 'formula', '=MIN(A1:A8)')
doc.set_cell_value(3, 3, 'formula', '=MAX(A1:A8)')
doc.set_cell_value(3, 4, 'formula', '=SUM(A1:A8)')
doc.set_cell_value(3, 5, 'formula', '=SQRT(A8)')
doc.set_cell_value(3, 6, 'formula', '=IF((A5>A4);A4;"")')

# Save the document to the file you want to create
doc.save("calc-example06.ods")
