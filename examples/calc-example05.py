#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create the document
doc = ooolib.Calc()

# Standard Cell Properties
# Cell Properties are handled by using on/off switches
# Turn the switch on, then do all the cells you want, then
# turn the switch back off.

doc.set_cell_value(1, 1, "string", "Normal Text")

doc.set_cell_property('bold', True)
doc.set_cell_value(1, 2, "string", "Bold Text")
doc.set_cell_property('bold', False)

doc.set_cell_property('italic', True)
doc.set_cell_value(1, 3, "string", "Italic Text")
doc.set_cell_property('italic', False)

doc.set_cell_property('underline', True)
doc.set_cell_value(1, 4, "string", "Underline Text")
doc.set_cell_property('underline', False)


# Colors
# Colors are in the format '#ffffff'.  If you use an
# incorrect format, the color will be changed to the
# default color.
doc.set_cell_property('color', '#0000ff')
doc.set_cell_property('background', '#ff0000')
doc.set_cell_value(2, 1, "string", "Blue on Red")

doc.set_cell_property('color', '#ff0000')
doc.set_cell_property('background', '#0000ff')
doc.set_cell_value(2, 2, "string", "Red on Blue")

doc.set_cell_property('color', 'default')
doc.set_cell_property('background', 'default')
doc.set_cell_value(2, 3, "string", "Default Colors")


# Text Font Sizes
doc.set_cell_property('fontsize', '10')
doc.set_cell_value(3, 1, "string", "Default 10pt")

doc.set_cell_property('fontsize', '11')
doc.set_cell_value(3, 2, "string", "11pt")

doc.set_cell_property('fontsize', '12')
doc.set_cell_value(3, 3, "string", "12pt")

doc.set_cell_property('fontsize', '10')

# Write out the document
doc.save("calc-example05.ods")
