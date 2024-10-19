import re
from enum import Enum, unique
from typing import Optional, Union
from xml.etree.ElementTree import Element

from .exceptions import CellPositionOutOfRange, InvalidCellPosition
from .mixin import RootMixin, attrsType, valueType


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

    def get_or_create_row(self, selected_row: int) -> Element:
        """Get or create row."""
        position = 0
        table = self.root_find("table:table")
        for table_row in table.findall('table:table-row', self.ns):
            repeated = table_row.get(self.qname("table:number-rows-repeated"))
            if repeated is None:
                position += 1
            else:
                position += int(repeated)
            if position == selected_row:
                return table_row
            if position > selected_row:
                break  # == find cell; > insert row before
        return self.get_or_create_element(table, "table:table-row")  # TODO:

    def get_or_create_cell(self, row: Element, selected_cell: int, attrs: Optional[attrsType] = None) -> Element:
        """Get or create cell."""
        return self.get_or_create_element(row, "table:table-cell", attrs)  # TODO:

    def set_cell_value(
            self, position: Union[str, tuple[int, int]],
            value: valueType,
            value_type: Optional[ValueType] = None
    ) -> None:
        """Set cell value."""
        selected_column, selected_row = self.get_coordinates(position)
        row = self.get_or_create_row(selected_row)
        if value_type is None:
            value_type = self.resolve_value_type(value)
        cell = self.get_or_create_cell(row, selected_column, {
            "office:value-type": value_type.value,
            "calcext:value-type": value_type.value,
        })
        self.get_or_create_element(cell, "text:p", value=str(value))
