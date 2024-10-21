import re
import xml.etree.ElementTree as ET
from enum import Enum, unique
from typing import Optional, Union
from xml.etree.ElementTree import Element

from .exceptions import CellPositionOutOfRange, InvalidCellPosition
from .mixin import RootMixin, attrsType, cast, valueType


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

    def count_rows(self) -> int:
        """Count table rows."""
        rows = 0
        for row in self.root.findall('table:table/table:table-row', self.ns):
            repeated = row.get(self.qname("table:number-rows-repeated"))
            if repeated is None:
                rows += 1
            else:
                rows += int(repeated)
        return rows

    def count_columns(self) -> int:
        """Count columns."""
        columns = 0
        for column in self.root.findall("table:table/table:table-column", self.ns):
            repeated = column.get(self.qname("table:number-columns-repeated"))
            if repeated is None:
                columns += 1
            else:
                columns += int(repeated)
        return columns

    def get_boundary(self) -> tuple[int, int]:
        """Get table max columns and rows."""
        return self.count_columns(), self.count_rows()

    def set_columns(self, columns: int) -> None:
        """Set columns."""
        column = self.root_find("table:table/table:table-column[last()]")
        repeated = column.get(self.qname("table:number-columns-repeated"))
        num = columns + 1 if repeated is None else int(repeated) + columns
        # column.set(self.qname("table:number-columns-repeated"), str(num))
        self.set_element_attr(column, "table:number-columns-repeated", num)

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

    def find_cells_or_covered(self, row: Element) -> list[Element]:
        """Find cells of covered cells."""
        # It would be better to use xpath query "table:table-cell|table:covered-table-cell",
        # but the ET has not yet implemented this feature.
        # Module lxml has it implemented, but there would be a dependency on it.
        cells = []
        for cell in row.findall("table:*", self.ns):
            if self.lname(cell.tag) in ("table-cell", "covered-table-cell"):
                cells.append(cell)
        return cells

    def append_rows(self, gap: int, table: Element, table_row: Element, columns: int) -> None:
        """Append rows with the attribute number-rows-repeated."""
        attrs = cast(attrsType, {}) if gap == 1 else {"table:number-rows-repeated": gap}
        cells = table_row.findall("table:*", self.ns)

        if (len(cells) == 1 and self.lname(cells[0].tag) in ("table-cell", "covered-table-cell")
                and cells[0].find("*") is None):
            # Row has only one cell and this cell does not have any child - so we can set attribute here.
            self.set_attrs(table_row, {"table:number-rows-repeated": gap})
        else:
            row = self.create_sub_element(table, "table:table-row", attrs)
            self.create_sub_element(row, "table:table-cell")

    # def append_columns(self, gap: int, table_row: Element) -> None:
    #     """Append columns with the attribute number-columns-repeated."""
    #     repeated = str(gap)
    #     attrs = {} if gap == 1 else {"table:number-columns-repeated": repeated}
    #     self.create_sub_element(table_row, "table:table-cell", attrs)

    def equal_table_element(
            self, position: int, parent: Element, child: Element, node_name: str, attr_name: str) -> Element:
        """Equal element."""
        attr_prefix_name = f"table:number-{attr_name}-repeated"
        node_prefix_name = f"table:table-{node_name}"
        repeated = child.get(self.qname(attr_prefix_name))
        print("repeated:", repeated)
        ET.dump(child)
        if repeated is None:
            return child
        self.set_attrs(child, {attr_prefix_name: int(repeated) - 1, "vlozeno": f"equal_{attr_name} repeated - 1"})
        row = self.create_element(node_prefix_name, {"vlozeno": f"equal_{node_name}", "position": position})
        parent.insert(position + 2, row)
        return row

    def equal_row(self, position: int, table: Element, table_row: Element) -> Element:
        """Equal row."""
        return self.equal_table_element(position, table, table_row, "row", "rows")
        # repeated = table_row.get(self.qname("table:number-rows-repeated"))
        # print("repeated:", repeated)
        # ET.dump(table_row)
        # if repeated is None:
        #     return table_row
        # self.set_attrs(table_row, {"table:number-rows-repeated": str(int(repeated) - 1),
        #                            "vlozeno": "equal_row repeated - 1"})
        # row = self.create_element("table:table-row", {"vlozeno": "equal_row", "position": str(position)})
        # table.insert(position + 2, row)
        # return row

    def equal_column(self, position: int, table_row: Element, table_cell: Element) -> Element:
        """Equal column."""
        return self.equal_table_element(position, table_row, table_cell, "cell", "columns")
        # repeated = table_cell.get(self.qname("table:number-columns-repeated"))
        # print("repeated:", repeated)
        # ET.dump(table_cell)
        # if repeated is None:
        #     return table_cell
        # self.set_attrs(table_cell, {"table:number-columns-repeated": str(int(repeated) - 1),
        #                            "vlozeno": "equal_column repeated - 1"})
        # row = self.create_element("table:table-cell", {"vlozeno": "equal_column", "position": str(position)})
        # table_row.insert(position + 2, row)
        # return row

    def insert_rows(
            self,
            position: int,
            rows_before: int,
            rows_after: int,
            selected_row: int,
            table: Element,
            table_row: Element,
            columns: int
    ) -> Element:
        """Insert rows."""
        ET.dump(table_row)
        print("position:", position, "rows_before:", rows_before, "rows_after:", rows_after,
              "selected_row:", selected_row)
        before = selected_row - rows_before - 1
        after = rows_after - selected_row
        print("before:", before, "after:", after)

        if before:
            attrs = cast(attrsType, {"vlozeno": "before == 1"}) if before == 1 else {
                "table:number-rows-repeated": before, "vlozeno": "before vetsi 1"}
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
                pass  # TODO: ???
            else:
                if after == 1:
                    self.pop_element_attr(table_row, "table:number-rows-repeated")
                else:
                    self.set_attrs(table_row, {"table:number-rows-repeated": after, "vlozeno": "after vetsi 1"})

        return row

    def insert_columns(
            self,
            position: int,
            columns_before: int,
            columns_after: int,
            selected_row: int,
            table_row: Element,
            table_cell: Element,
    ) -> Element:
        """Append rows."""
        ET.dump(table_row)
        ET.dump(table_cell)  # TODO: použít nebo rozdělit
        print("position:", position, "rows_after:", columns_after, "selected_row:", selected_row)
        before = selected_row - columns_before - 1
        after = columns_after - selected_row
        print("before:", before, "after:", after)
        repeated = self.get_element_attr(table_cell, "table:number-columns-repeated", 0)
        # value = table_cell.get(self.qname("table:number-columns-repeated"))
        # if value is None:
        #     repeated = 0
        # else:
        #     repeated = int(value)
        print("repeated:", repeated)
        # import pdb; pdb.set_trace()  # !!!
        if before:
            attrs = cast(attrsType, {}) if before == 1 else {"table:number-columns-repeated": before}
            cell_before = self.create_element("table:table-cell", attrs)
            table_row.insert(position, cell_before)

        if after:
            attrs = cast(attrsType, {}) if after == 1 else {"table:number-columns-repeated": after}
            cell_after = self.create_element("table:table-cell", attrs)
            table_row.insert(position + 1, cell_after)

        self.pop_element_attr(table_cell, "table:number-columns-repeated")

        # if before:
        #     attrs = {"vlozeno": "before == 1"} if before == 1 else {"table:number-rows-repeated": str(before),
        #                                                             "vlozeno": "before vetsi 1"}
        #     row_before = self.create_element("table:table-row", attrs)
        #     self.create_sub_element(row_before, "table:table-cell")
        #     position += 1
        #     table.insert(position, row_before)

        # row = self.create_element("table:table-row", {"vlozeno": "hodnota"})
        # position += 1
        # table.insert(position, row)

        # if after:
        #     repeated = table_row.get(self.qname("table:number-rows-repeated"))
        #     print("repeated:", repeated)
        #     if repeated is None:
        #         pass
        #     else:
        #         if after == 1:
        #             table_row.attrib.pop(self.qname("table:number-rows-repeated"))
        #         else:
        #             self.set_attrs(table_row, {"table:number-rows-repeated": str(after), "vlozeno": "after vetsi 1"})

        return table_cell

    def get_or_create_row(self, selected_row: int, columns: int) -> Element:
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
                print("=== ROW EQUAL", "=" * 60)
                return self.equal_row(position, table, table_row)
            if rows_after > selected_row:
                print(">" * 60)
                return self.insert_rows(position, rows_before, rows_after, selected_row, table, table_row, columns)
            if repeated is None:
                rows_before += 1
            else:
                rows_before += int(repeated)
        gap = selected_row - rows_after - 1
        print("### ROW gap:", gap)
        if gap > 0:
            self.append_rows(gap, table, table_row, columns)
        return self.create_sub_element(table, "table:table-row")

    def get_or_create_cell(
            self,
            row: Element,
            selected_column: int,
            value: valueType,
            value_type: ValueType,
            columns: int
    ) -> Element:
        """Get or create cell."""
        # cell = self.get_or_create_element(row, "table:table-cell", attrs, alternate_name="table:covered-table-cell")
        print("c" * 60)

        columns_before, columns_after = 0, 0
        for position, table_cell in enumerate(row.findall('table:*', self.ns)):
            if self.lname(table_cell.tag) not in ("table-cell", "covered-table-cell"):
                continue
            repeated = table_cell.get(self.qname("table:number-columns-repeated"))
            # print("repeated:", repeated)
            if repeated is None:
                columns_after += 1
            else:
                columns_after += int(repeated)
            # print("position:", position, "rowsnum:", rowsnum, "selected_row:", selected_row)
            if columns_after == selected_column:
                print("=== COLUMN EQUAL", "=" * 60)
                return self.equal_column(position, row, table_cell)
                # return table_cell
            if columns_after > selected_column:
                print(">>> COLUMN ", ">" * 60)
                return self.insert_columns(position, columns_before, columns_after, selected_column, row, table_cell)
            if repeated is None:
                columns_before += 1
            else:
                columns_before += int(repeated)

        # import pdb; pdb.set_trace()  # !!!

        gap = selected_column - columns_before
        print("### BEFORE gap:", gap)
        if gap > 1:
            repeated = table_cell.get(self.qname("table:number-columns-repeated"))
            if repeated is None:
                if table_cell.find("*") is None:
                    # table_cell.set(self.qname("table:number-columns-repeated"), str(gap))
                    self.set_element_attr(table_cell, "table:number-columns-repeated", gap)
                else:
                    self.create_sub_element(row, "table:table-cell", {"table:number-columns-repeated": gap})
            else:
                # table_cell.set(self.qname("table:number-columns-repeated"), str(gap + int(repeated)))
                self.set_element_attr(table_cell, "table:number-columns-repeated", gap + int(repeated))

        attrs = cast(attrsType, {
            "office:value-type": value_type.value,
            "calcext:value-type": value_type.value,
        })
        if value_type == ValueType.float:
            attrs["office:value"] = value

        cell = self.create_sub_element(row, "table:table-cell", attrs)

        if value_type != ValueType.float:
            self.pop_element_attr(cell, "office:value")

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
        print("selected_column, selected_row:", selected_column, selected_row)

        columns = self.count_columns()
        print("GET:columns:", columns, "selected_column:", selected_column)
        if selected_column > columns:
            self.set_columns(selected_column - columns)
            columns = selected_column
        print("columns:", columns)

        row = self.get_or_create_row(selected_row, columns)

        if value_type is None:
            value_type = self.resolve_value_type(value)

        cell = self.get_or_create_cell(row, selected_column, value, value_type, columns)
        self.get_or_create_element(cell, "text:p", value=value)

        ET.indent(self.root)
        ET.dump(self.root)
