#!/usr/bin/env python3.9
from ooolib.document import Calc

calc = Calc()
sheet = calc.get_sheet()
print("sheet:", sheet)
# calc.load("test-01.ods")
calc.save("test-11-out.ods")
