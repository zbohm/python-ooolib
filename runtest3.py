#!/usr/bin/env python3.9
import xml.etree.ElementTree as ET
from ooolib.document import Calc

calc = Calc()

sheet = calc.get_sheet()
# print("Boundary:", sheet.get_boundary())

sheet.set_cell_value("A1", "one")

calc.save("/tmp/devel.ods")
# ET.dump(sheet.root)
# ./runtest3.py | xmllint --encode utf8 --format -

# calc.debug_cells()
# sheet = calc.get_sheet(1)
# ET.dump(sheet.root)
