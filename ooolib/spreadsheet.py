import re
import xml.etree.ElementTree as ET
from enum import Enum, unique
from typing import Optional, Union
from xml.etree.ElementTree import Element

from .exceptions import CellPositionOutOfRange, InvalidCellPosition
from .mixin import RootMixin


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

    def resolve_value_type(self, value: Union[str, int, float]) -> ValueType:
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

    def find_row(self, selected_row: int) -> tuple[int, Optional[Element]]:
        """Find row."""
        position_row, table_row = 0, None
        for table_row in self.root.findall('table:table/table:table-row', self.ns):
            repeated = table_row.get(self.qname("table:number-rows-repeated"))
            if repeated is None:
                position_row += 1
            else:
                position_row += int(repeated)
            if position_row >= selected_row:
                break  # == find cell; > insert row before
        return position_row, table_row

    def create_row(self) -> Element:
        """Create row."""
        table = self.root_find('table:table')
        return ET.SubElement(table, self.qname("table:table-row"))

    def create_cell(self, parent: Element, attrs: Optional[dict[str, str]] = None) -> Element:
        """Create cell."""
        if attrs is None:
            attrs = {}
        return ET.SubElement(parent, self.qname("table:table-cell"), attrs)

    def create_text_p(self, parent: Element, value: str):
        """Create text paragraph."""
        node = ET.SubElement(parent, self.qname("text:p"))
        node.text = value
        return node

    def set_cell_value(
            self, position: Union[str, tuple[int, int]],
            value: Union[str, int, float],
            value_type: Optional[ValueType] = None
    ) -> None:
        """Set cell value."""
        selected_column, selected_row = self.get_coordinates(position)
        if value_type is None:
            value_type = self.resolve_value_type(value)
        position_row, table_row = self.find_row(selected_row)
        if table_row is None:
            element_row = self.create_row()
            cell = self.create_cell(element_row, {self.qname("office:value-type"): value_type.value})
            self.create_text_p(cell, str(value))
