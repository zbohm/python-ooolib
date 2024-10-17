#!/usr/bin/env python3.9
import xml.etree.ElementTree as ET
from ooolib.document import Calc

calc = Calc()

sheet = calc.get_sheet()
# print("Boundary:", sheet.boundary)

sheet.set_cell_value("A1", "test")

ET.dump(sheet.root)
# ./runtest3.py | xmllint --encode utf8 --format -
