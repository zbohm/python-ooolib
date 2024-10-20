import re
import xml.etree.ElementTree as ET
from enum import Enum, unique
from typing import Optional, Union
from xml.etree.ElementTree import Element

from .exceptions import CellPositionOutOfRange, InvalidCellPosition
from .mixin import RootMixin, valueType


@unique
class ValueType(Enum):
    string = "string"
    float = "float"
    formula = "formula"
    annotation = "annotation"
    link = "link"


class Spreadsheet(RootMixin):
    """Calc Spreadsheet."""

    def __init__(self, sheet: Element) -> None:
        super().__init__()
        self.root = sheet
        self.boundary: tuple[int, int] = self.calculate_boundary()

    def calculate_boundary(self) -> tuple[int, int]:
        """Calculate table max rows and columns."""
        rows, columns = 0, 0
        count_columns = True
        for row in self.root.findall('table:table/table:table-row', self.ns):
            repeated = row.get(self.qname("table:number-rows-repeated"))
            if repeated is None:
                rows += 1
            else:
                rows += int(repeated)
            if count_columns:
                count_columns = False  # count only once
                for cell in row.findall("table:table-cell", self.ns):
                    repeated = cell.get(self.qname("table:number-columns-repeated"))
                    if repeated is None:
                        spanned = cell.get(self.qname("table:number-columns-spanned"))
                        if spanned is None:
                            columns += 1
                        else:
                            columns += int(spanned)
                    else:
                        columns += int(repeated)
        return rows, columns

    def resolve_value_type(self, value: valueType) -> ValueType:
        """Resolve value type."""
        value_type = ValueType.string
        if isinstance(value, (int, float)):
            value_type = ValueType.float
        elif isinstance(value, str):
            if value[:1] == "=":
                value_type = ValueType.formula
            elif re.match("(https?|file)://", value):
                value_type = ValueType.link
        return value_type

    def get_coordinates(self, position: Union[str, tuple[int, int]]) -> tuple[int, int]:
        """Get coordinates."""
        if isinstance(position, tuple):
            column, row = position
        else:
            match = re.match(r"(?P<column>[A-Z]{1,3})(?P<row>\d+)$", position)
            if match is None:
                raise InvalidCellPosition(position)
            scolumn, srow = match.group("column"), match.group("row")
            column = 0
            for c in scolumn:
                column = column * 26 + (ord(c) - 65) + 1
            row = int(srow)
        # https://wiki.documentfoundation.org/Faq/Calc/022
        if not (0 < column < 0x4000 and 0 < row < 2**20):
            raise CellPositionOutOfRange(position)
        return column, row

    def append_rows(self, gap: int, table: Element, table_row: Element) -> None:
        """Append rows with the attribute number-rows-repeated."""
        repeated = str(gap)
        attrs = {} if gap == 1 else {"table:number-rows-repeated": repeated}
        cells = table_row.findall("table:*", self.ns)

        if (len(cells) == 1 and self.lname(cells[0].tag) in ("table-cell", "covered-table-cell")
                and cells[0].find("*") is None):
            # Row has only one cell and this cell does not have any child - so we can set attribute here.
            self.set_attrs(table_row, {"table:number-rows-repeated": repeated})
        else:
            row = self.create_sub_element(table, "table:table-row", attrs)
            self.create_sub_element(row, "table:table-cell")

    def equal_row(self, position: int, table: Element, table_row: Element) -> Element:
        """Equal row."""
        repeated = table_row.get(self.qname("table:number-rows-repeated"))
        print("repeated:", repeated)
        ET.dump(table_row)
        if repeated is None:
            return table_row
        self.set_attrs(table_row, {"table:number-rows-repeated": str(int(repeated) - 1),
                                   "vlozeno": "equal_row repeated - 1"})
        row = self.create_element("table:table-row", {"vlozeno": "equal_row", "position": str(position)})
        table.insert(position + 2, row)
        return row

    def insert_rows(
            self,
            position: int,
            rows_before: int,
            rows_after: int,
            selected_row: int,
            table: Element,
            table_row: Element
    ) -> Element:
        """Append rows."""
        ET.dump(table_row)
        print("position:", position, "rows_before:", rows_before, "rows_after:", rows_after,
              "selected_row:", selected_row)
        before = selected_row - rows_before - 1
        after = rows_after - selected_row
        print("before:", before, "after:", after)

        if before:
            attrs = {"vlozeno": "before == 1"} if before == 1 else {"table:number-rows-repeated": str(before),
                                                                    "vlozeno": "before vetsi 1"}
            row_before = self.create_element("table:table-row", attrs)
            self.create_sub_element(row_before, "table:table-cell")
            position += 1
            table.insert(position, row_before)

        row = self.create_element("table:table-row", {"vlozeno": "hodnota"})
        position += 1
        table.insert(position, row)

        if after:
            repeated = table_row.get(self.qname("table:number-rows-repeated"))
            print("repeated:", repeated)
            if repeated is None:
                pass
            else:
                if after == 1:
                    table_row.attrib.pop(self.qname("table:number-rows-repeated"))
                else:
                    self.set_attrs(table_row, {"table:number-rows-repeated": str(after), "vlozeno": "after vetsi 1"})

        return row

    def get_or_create_row(self, selected_row: int) -> Element:
        """Get or create row."""
        rows_before, rows_after = 0, 0
        table = self.root_find("table:table")
        for position, table_row in enumerate(table.findall('table:table-row', self.ns)):
            repeated = table_row.get(self.qname("table:number-rows-repeated"))
            # print("repeated:", repeated)
            if repeated is None:
                rows_after += 1
            else:
                rows_after += int(repeated)
            # print("position:", position, "rowsnum:", rowsnum, "selected_row:", selected_row)
            if rows_after == selected_row:
                print("=== EQUAL", "=" * 60)
                return self.equal_row(position, table, table_row)
            if rows_after > selected_row:
                print(">" * 60)
                return self.insert_rows(position, rows_before, rows_after, selected_row, table, table_row)
            if repeated is None:
                rows_before += 1
            else:
                rows_before += int(repeated)
        gap = selected_row - rows_after - 1
        print("### gap:", gap)
        if gap > 0:
            self.append_rows(gap, table, table_row)
        return self.create_sub_element(table, "table:table-row")

    def get_or_create_cell(
            self,
            row: Element,
            selected_column: int,
            value: valueType,
            value_type: ValueType
    ) -> Element:
        """Get or create cell."""
        attrs = {
            "office:value-type": value_type.value,
            "calcext:value-type": value_type.value,
        }
        if value_type == ValueType.float:
            attrs["office:value"] = str(value)
        cell = self.get_or_create_element(row, "table:table-cell", attrs)
        if value_type != ValueType.float:
            cell.attrib.pop(self.qname("office:value"))
        return cell

    def set_cell_value(
            self, position: Union[str, tuple[int, int]],
            value: valueType,
            value_type: Optional[ValueType] = None
    ) -> None:
        """Set cell value."""
        ET.indent(self.root)
        ET.dump(self.root)
        print("." * 60)
        selected_column, selected_row = self.get_coordinates(position)

        row = self.get_or_create_row(selected_row)

        if value_type is None:
            value_type = self.resolve_value_type(value)

        cell = self.get_or_create_cell(row, selected_column, value, value_type)
        self.get_or_create_element(cell, "text:p", value=value)

        ET.indent(self.root)
        ET.dump(self.root)
