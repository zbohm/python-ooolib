#!/usr/bin/env python3.9
from ooolib.document import Calc


calc = Calc()

calc.load("F10.ods")

calc.debug_cells()

sheet = calc.get_sheet()
print("Boundary:", sheet.boundary)

# sheet.set_cell_value(2, 3, "test")
# sheet.set_cell_value(3, 3, 42)
# sheet.set_cell_value(4, 3, 3.12)
# sheet.set_cell_value(5, 3, "=SUM()")
# sheet.set_cell_value(6, 3, "https://example.com/")

name = "F10-out.ods"
calc.save(name)
print(f"Saved file {name}")

calc2 = Calc()
calc2.load(name)
sheet = calc2.get_sheet()
print("Boundary2:", sheet.boundary)
