#!/usr/bin/env python3.9
from ooolib.document import Calc


calc = Calc()

calc.load("F10.ods")

calc.debug_cells()
sheet = calc.get_sheet()
print("Boundary:", sheet.get_boundary())

# sheet.set_cell_value((2, 3), "test")

# sheet.set_cell_value("A1", "test")
# # sheet.set_cell_value("A2", "test")
# sheet.set_cell_value("Z1", "test")
# sheet.set_cell_value("AA1", "test")
# print("701:", end="")
# sheet.set_cell_value("ZZ1", "test")
# print("702:", end="")
# sheet.set_cell_value("AAA1", "test")

# sheet.set_cell_value("AZ1", "test")
# sheet.set_cell_value("ZZ1", "test")
# sheet.set_cell_value("XFD1", "max")
# sheet.set_cell_value("XFE1", "over")
# sheet.set_cell_value("ZZZ1", "test")

print("  A  1 ", end=""); sheet.set_cell_value("A1", "A1")
print("  Z 26 ", end=""); sheet.set_cell_value("Z1", "Z1")
print(" AA 27 ", end=""); sheet.set_cell_value("AA1", "AA1")
print(" AZ 52", end=""); sheet.set_cell_value("AZ1", "AZ1")
print(" BA 53", end=""); sheet.set_cell_value("BA1", "BA1")
print(" BB 54", end=""); sheet.set_cell_value("BB1", "BB1")
print(" BZ 78", end=""); sheet.set_cell_value("BZ1", "BZ1")
print(" CA 79", end=""); sheet.set_cell_value("CA1", "CA1")
print(" ZA 677", end=""); sheet.set_cell_value("ZA1", "ZA1")
print(" ZB 678", end=""); sheet.set_cell_value("ZB1", "ZB1")
print(" ZZ 702", end=""); sheet.set_cell_value("ZZ1", "ZZ1")
print("AAA 703", end=""); sheet.set_cell_value("AAA1", "AAA1")
print("AAB 704", end=""); sheet.set_cell_value("AAB1", "AAB1")
sheet.set_cell_value("XFC1", "XFC1")
# sheet.set_cell_value("XFE1", "XFE1")
# sheet.set_cell_value("ZZZ1", "ZZZ1")


# sheet.set_cell_value(2, 3, "test")
# sheet.set_cell_value(3, 3, 42)
# sheet.set_cell_value(4, 3, 3.12)
# sheet.set_cell_value(5, 3, "=SUM()")
# sheet.set_cell_value(6, 3, "https://example.com/")

# name = "F10-out.ods"
# calc.save(name)
# print(f"Saved file {name}")

# calc2 = Calc()
# calc2.load(name)
# sheet = calc2.get_sheet()
# print("Boundary2:", sheet.get_boundary())
