import re
from enum import Enum, unique
from typing import Union
from xml.etree.ElementTree import Element

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
        self.boundary: list[int, int] = self.calculate_boundary()

    def calculate_boundary(self) -> list[int, int]:
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

    def set_cell_value(
            self, row: int, column: int, value: Union[str, int, float], value_type: ValueType = None
    ) -> None:
        """Set cell value."""
        print("=-" * 60)
        print(row, column)
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

"""
10 x 6

       A       B        C      D       E        F
   +-------+
1  |  A1   |
   +-------+-------+
2  |       |   22  |
   +-------+-------+
3  |       |=SUM(a)|
   +-------+-------+
4  |       |    (a)|
   +-------+-------+-------+-------+-------+
5  |       |       |  TU           |   E5  |
   +-------+-------+-------+-------+-------+
6  |       |       |       |       |       |
   +-------+-------+-------+-------+-------+
7  |       |       |               |       |
   +-------+-------+               +-------+
8  |       |       |               |       |
   +-------+-------+               +-------+
9  |       |       |  SOM          |       |
   +-------+-------+-------+-------+-------+-------+
10 | http..|       |       |       |       |  F10  |
   +-------+-------+-------+-------+-------+-------+

<office:spreadsheet>
      <table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>
      <table:table table:name="List1" table:style-name="ta1">
        <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
        <table:table-column table:style-name="co2" table:number-columns-repeated="5" table:default-cell-style-name="Default"/>

1       <table:table-row table:style-name="ro1">
          <table:table-cell office:value-type="string" calcext:value-type="string">
            <text:p>A1</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="5"/>
        </table:table-row>

2       <table:table-row table:style-name="ro1">
          <table:table-cell/>
          <table:table-cell office:value-type="float" office:value="22" calcext:value-type="float">
            <text:p>22</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="4"/>
        </table:table-row>

3       <table:table-row table:style-name="ro1">
          <table:table-cell/>
          <table:table-cell table:formula="of:=SUM([.B2:.B2])" office:value-type="float" office:value="22" calcext:value-type="float">
            <office:annotation draw:style-name="gr1" draw:text-style-name="P2" svg:width="2.899cm" svg:height="1.799cm" svg:x="6.283cm" svg:y="0cm" draw:caption-point-x="-0.61cm" draw:caption-point-y="0.913cm">
              <dc:date>2024-10-16T00:00:00</dc:date>
              <text:p text:style-name="P1">
                <text:span text:style-name="T1">Summary.</text:span>
              </text:p>
            </office:annotation>
            <text:p>22</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="4"/>
        </table:table-row>

4       <table:table-row table:style-name="ro1">
          <table:table-cell/>
          <table:table-cell>
            <office:annotation draw:style-name="gr1" draw:text-style-name="P2" svg:width="2.899cm" svg:height="1.799cm" svg:x="6.283cm" svg:y="0cm" draw:caption-point-x="-0.61cm" draw:caption-point-y="1.364cm">
              <dc:date>2024-10-16T00:00:00</dc:date>
              <text:p text:style-name="P1">
                <text:span text:style-name="T1">Empty.</text:span>
              </text:p>
            </office:annotation>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="4"/>
        </table:table-row>

5       <table:table-row table:style-name="ro1">
          <table:table-cell table:number-columns-repeated="2"/>
          <table:table-cell office:value-type="string" calcext:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="1">
            <text:p>TU</text:p>
          </table:table-cell>
          <table:covered-table-cell/>
          <table:table-cell office:value-type="string" calcext:value-type="string">
            <text:p>E5</text:p>
          </table:table-cell>
          <table:table-cell/>
        </table:table-row>

6       <table:table-row table:style-name="ro1">
          <table:table-cell table:number-columns-repeated="6"/>
        </table:table-row>

7       <table:table-row table:style-name="ro1">
          <table:table-cell table:number-columns-repeated="2"/>
          <table:table-cell office:value-type="string" calcext:value-type="string" table:number-columns-spanned="1" table:number-rows-spanned="3">
            <text:p>SOM</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="3"/>
        </table:table-row>

8       <table:table-row table:style-name="ro1">
          <table:table-cell table:number-columns-repeated="2"/>
          <table:covered-table-cell/>
          <table:table-cell table:number-columns-repeated="3"/>
        </table:table-row>

9       <table:table-row table:style-name="ro1">
          <table:table-cell table:number-columns-repeated="2"/>
          <table:covered-table-cell/>
          <table:table-cell table:number-columns-repeated="3"/>
        </table:table-row>

10     <table:table-row table:style-name="ro1">
          <table:table-cell office:value-type="string" calcext:value-type="string">
            <text:p>
              <text:a xlink:href="https://example.com/" xlink:type="simple">https://example.com</text:a>
            </text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="4"/>
          <table:table-cell office:value-type="string" calcext:value-type="string">
            <text:p>F10</text:p>
          </table:table-cell>
        </table:table-row>

      </table:table>
      <table:named-expressions/>
    </office:spreadsheet>
"""