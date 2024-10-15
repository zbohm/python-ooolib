#/ust/bin/python
import ooolib
import re

src_file = "decka.ods"
out_file = "zkouska.ods"


doc_src = ooolib.Calc(opendoc=src_file)
doc_src.set_sheet_index(0)
(cols, rows) = doc_src.get_sheet_dimensions()
print("cols, rows:", cols, rows)

doc_dest = ooolib.Calc(opendoc=out_file)
doc_dest.set_sheet_index(0)

for r in range(rows + 1):
    for c in range(cols + 1):
        print("c, r:", c, r)
        if cell_value := doc_src.get_cell_value(c, r):
            print("cell_value:", cell_value)
            value_type, value = cell_value
            #  set_cell_value(self, col, row, datatype, value, formula_value='0'):
            if value_type == "formula":
                print("??? value:", value)
                value = re.sub("^of:", "", value, 1)
                value = re.sub("\.([A-Z])", r"\1", value)
                value = re.sub(r"\(\[(.+?)\]\)", r"(\1)", value)
                # 'of:=SUM([.A5:.A6])')
                print("!!! value:", value)
            doc_dest.set_cell_value(c, r, value_type, value)

doc_dest.save(out_file)
