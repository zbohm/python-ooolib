import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .content import Content


class Sheet(Content):

    default_list_name = "List"

    def create(self) -> Element:
        """Create content."""
        root = ET.Element(self.qualify('office:document-content'), {
            "xmlns:calcext": self.ns["calcext"],
            "office:version": self.version,
        })
        ET.SubElement(root, self.qualify('office:body'))
        self.create_default_sheet(root)
        return root

    def create_default_sheet(self, root: Element, name: str = None) -> Element:
        """Create default sheet."""
        body = root.find("office:body", self.ns)
        sheets = body.findall("office:spreadsheet", self.ns)
        if name is None:
            name = f"{self.default_list_name}{len(sheets) + 1}"
        sheet = ET.SubElement(body, self.qualify('office:spreadsheet'))
        ET.SubElement(sheet, self.qualify('table:table'), {"table:name": name})
        return sheet

    def create_sheet(self, name: str = None) -> Element:
        """Create sheet."""
        return self.create_default_sheet(self.get_or_create_root(), name)

    def debug_cells(self) -> None:
        """Debug cells."""
        sheet = self.create_sheet()
        table = sheet.find("table:table", self.ns)
        row = ET.SubElement(table, self.qualify('table:table-row'))
        cell = ET.SubElement(row, self.qualify('table:table-cell'), {
            "office:value-type": "float",
            "office:value": "1",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, self.qualify('text:p'))
        text.text = "1"
        cell = ET.SubElement(row, self.qualify('table:table-cell'), {
            "office:value-type": "float",
            "office:value": "2",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, self.qualify('text:p'))
        text.text = "2"
        cell = ET.SubElement(row, self.qualify('table:table-cell'), {
            "office:value-type": "float",
            "office:value": "3",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, self.qualify('text:p'))
        text.text = "3"

        cell = ET.SubElement(row, self.qualify('table:table-cell'), {
            "office:value-type": "string",
            "calcext:value-type": "string",
        })
        text = ET.SubElement(cell, self.qualify('text:p'))
        text.text = "Matěj"
