from typing import Optional
from xml.etree.ElementTree import Element

from .content import Content
from .exceptions import ElementNotFound
from .spreadsheet import Spreadsheet


class Calc(Content):

    default_list_name = "List"

    def create(self) -> Element:
        """Create content."""
        root = self.create_element("office:document-content", {
            "xmlns:calcext": self.ns["calcext"],
            "office:version": self.version
        })
        self.create_default_sheet(root)
        return root

    def create_automatic_styles(self, root: Element) -> None:
        """Create automatic styles."""
        automatic_styles = self.create_sub_element(root, 'office:automatic-styles')
        style = self.create_sub_element(automatic_styles, 'style:style', {
            "style:name": "co1",
            "style:family": "table-column",
        })
        self.create_sub_element(style, 'style:table-column-properties', {
            "fo:break-before": "auto",
            "style:column-width": "2.258cm",
        })
        style = self.create_sub_element(automatic_styles, 'style:style', {
            "style:name": "ro1",
            "style:family": "table-row",
        })
        self.create_sub_element(style, 'style:table-row-properties', {
            "style:row-height": "0.452cm",
            "fo:break-before": "auto",
            "style:use-optimal-row-height": "true",
        })
        style = self.create_sub_element(automatic_styles, 'style:style', {
            "style:name": "ta1",
            "style:family": "table",
            "style:master-page-name": "Default",
        })
        self.create_sub_element(style, 'style:table-properties', {
            "table:display": "true",
            "style:writing-mode": "lr-tb",
        })

    def create_default_sheet(self, root: Element, name: Optional[str] = None) -> Element:
        """Create default sheet."""
        self.create_automatic_styles(root)
        body = self.create_sub_element(root, 'office:body')
        sheets = body.findall("office:spreadsheet", self.ns)
        if name is None:
            name = f"{self.default_list_name}{len(sheets) + 1}"
        sheet = self.create_sub_element(body, 'office:spreadsheet')
        self.create_sub_element(sheet, 'table:calculation-settings', {
            "table:automatic-find-labels": "false",
            "table:use-regular-expressions": "false",
            "table:use-wildcards": "true",
        })
        table = self.create_sub_element(sheet, 'table:table', {
            "table:name": name,
            "table:style-name": "ta1",
        })
        self.create_sub_element(sheet, "table:named-expressions")
        self.create_sub_element(table, 'table:table-column', {
            "table:style-name": "co1", "table:default-cell-style-name": "Default",
        })
        row = self.create_sub_element(table, 'table:table-row', {"table:style-name": "ro1"})
        self.create_sub_element(row, 'table:table-cell')
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
