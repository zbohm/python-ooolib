#!/usr/bin/env python3.9
from pathlib import Path
from ooolib.document import Calc


calc = Calc()
# calc.load(Path().home().joinpath('tests').joinpath('calc-cells.ods'))
calc.load(Path().home().joinpath('tests').joinpath('calc-covered-cells.ods'))
# calc.load(Path().home().joinpath('tests').joinpath('c1-3.ods'))

sheet = calc.get_sheet()
# sheet.set_cell_value("A1", "one")
# sheet.set_cell_value("A2", 2)
# sheet.set_cell_value("A3", 3)
# sheet.set_cell_value("A4", 4)
sheet.set_cell_value("A5", 5)
# sheet.set_cell_value("A7", 7)
# sheet.set_cell_value("A6", 6)
# sheet.set_cell_value("A8", 8)
# sheet.set_cell_value("A9", 9)
sheet.set_cell_value("A10", "ten")
# sheet.set_cell_value("A11", 11)
# sheet.set_cell_value("A12", 12)
# sheet.set_cell_value("A13", 13)
# sheet.set_cell_value("A20", 20)

calc.save(Path().home().joinpath('tests').joinpath('debug.ods'))
