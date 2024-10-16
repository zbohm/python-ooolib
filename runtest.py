#!/usr/bin/env python3.9
from ooolib.document import Calc


calc = Calc()

calc.load("Iva-20.ods")

# calc.debug_cells()
sheet = calc.get_sheet()
print("Boundary:", sheet.boundary)

sheet.set_cell_value(2, 3, "test")
# sheet.set_cell_value(3, 3, 42)
# sheet.set_cell_value(4, 3, 3.12)
# sheet.set_cell_value(5, 3, "=SUM()")
# sheet.set_cell_value(6, 3, "https://example.com/")

# calc.load("test-01.ods")
# calc.save("test-11-out.ods")
