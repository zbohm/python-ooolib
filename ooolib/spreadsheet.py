from xml.etree.ElementTree import Element

from .mixin import RootMixin


class Spreadsheet(RootMixin):
    """Calc Spreadsheet."""

    def __init__(self, sheet: Element) -> None:
        super().__init__()
        self.root = sheet

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
                        columns += 1
                    else:
                        columns += int(repeated)
        return rows, columns
