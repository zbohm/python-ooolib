import xml.etree.ElementTree as ET
from typing import Optional
from xml.etree.ElementTree import Element

from .content import Content
from .exceptions import ElementNotFound
from .spreadsheet import Spreadsheet


class Sheet(Content):

    default_list_name = "List"

    def create(self) -> Element:
        """Create content."""
        root = ET.Element(self.qname('office:document-content'), {
            "xmlns:calcext": self.ns["calcext"],
            "office:version": self.version,
        })
        ET.SubElement(root, self.qname('office:body'))
        self.create_default_sheet(root)
        return root

    def create_default_sheet(self, root: Element, name: Optional[str] = None) -> Element:
        """Create default sheet."""
        body = root.find("office:body", self.ns)
        if body is None:
            raise ElementNotFound("office:body")
        sheets = body.findall("office:spreadsheet", self.ns)
        if name is None:
            name = f"{self.default_list_name}{len(sheets) + 1}"
        sheet = ET.SubElement(body, self.qname('office:spreadsheet'))
        ET.SubElement(sheet, self.qname('table:table'), {self.qname("table:name"): name})
        return sheet

    def create_sheet(self, name: Optional[str] = None) -> Spreadsheet:
        """Create sheet."""
        return Spreadsheet(self.create_default_sheet(self.get_or_create_root(), name))

    def get_sheet(self, position: int = 0) -> Spreadsheet:
        """Get sheet."""
        root = self.get_or_create_root()
        sheets = root.findall("office:body/office:spreadsheet", self.ns)
        return Spreadsheet(sheets[position])

    def debug_cells(self) -> None:
        """Debug cells."""
        sheet = self.create_sheet()
        table = sheet.root.find("table:table", self.ns)
        if table is None:
            raise ElementNotFound("table:table")

        row = self.create_sub_element(table, "table:table-row")

        cell = self.create_sub_element(row, "table:table-cell", {
            "office:value-type": "float",
            "office:value": "1",
            "calcext:value-type": "float",
        })
        self.create_sub_element(cell, "text:p", value="1")

        cell = self.create_sub_element(row, "table:table-cell", {
            "office:value-type": "float",
            "office:value": "2",
            "calcext:value-type": "float",
        })
        self.create_sub_element(cell, "text:p", value="2")

        cell = self.create_sub_element(row, "table:table-cell", {
            "office:value-type": "float",
            "office:value": "3",
            "calcext:value-type": "float",
        })
        self.create_sub_element(cell, "text:p", value="3")

        cell = self.create_sub_element(row, "table:table-cell", {
            "office:value-type": "string",
            "calcext:value-type": "string",
        })
        self.create_sub_element(cell, "text:p", value="Matěj")
