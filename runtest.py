#!/usr/bin/env python3.9
from ooolib.document import Calc


calc = Calc()

calc.load("Iva-20.ods")

# calc.debug_cells()
sheet = calc.get_sheet()
print("Boundary:", sheet.calculate_boundary())

# calc.load("test-01.ods")
# calc.save("test-11-out.ods")
