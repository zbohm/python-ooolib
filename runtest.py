#!/usr/bin/env python3.9
from ooolib.document import Calc

calc = Calc()
# calc.load("test-01.ods")
# calc.save("test-10-out.ods")

sheet = calc.get_sheet()
print("sheet:", sheet)
