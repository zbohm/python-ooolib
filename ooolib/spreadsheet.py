import re
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
            repeated = row.get(self.qualify("table:number-rows-repeated"))
            if repeated is None:
                rows += 1
            else:
                rows += int(repeated)
            if count_columns:
                count_columns = False  # count only once
                for cell in row.findall("table:table-cell", self.ns):
                    repeated = cell.get(self.qualify("table:number-columns-repeated"))
                    if repeated is None:
                        spanned = cell.get(self.qualify("table:number-columns-spanned"))
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
            elif re.match("https?://", value):
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
            num = 0
            for c in scolumn:
                num = num * 26 + (ord(c) - 65) + 1
            column = num - 1
            row = int(srow) - 1
        # https://wiki.documentfoundation.org/Faq/Calc/022
        if not (-1 < column < 0x4000 and -1 < row < 2**20):
            raise CellPositionOutOfRange(position)
        return column, row

    def set_cell_value(
            self, position: Union[str, tuple[int, int]],
            value: Union[str, int, float],
            value_type: Optional[ValueType] = None
    ) -> None:
        """Set cell value."""
        column, row = self.get_coordinates(position)
        print(":", position, column, row)
        if value_type is None:
            value_type = self.resolve_value_type(value)
        print("value:", repr(value), "value_type:", value_type)
        position_row = 0
        for table_row in self.root.findall('table:table/table:table-row', self.ns):
            if row == position_row:
                break
            repeated = table_row.get(self.qualify("table:number-rows-repeated"))
            if repeated is None:
                position_row += 1
            else:
                position_row += int(repeated)
            print("position_row:", position_row)
