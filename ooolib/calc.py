import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import RootMixin


class Sheet(RootMixin):

    filename = "content.xml"

    def create(self) -> Element:
        """Create content."""
        root = ET.Element('office:document-content', {
            "xmlns:office": self.ns["office"],
            "xmlns:table": self.ns["table"],
            "xmlns:text": self.ns["text"],
            "xmlns:calcext": self.ns["calcext"],
            "office:version": self.version,
        })
        body = ET.SubElement(root, 'office:body')
        sheet = ET.SubElement(body, 'office:spreadsheet')
        ET.SubElement(sheet, 'table:table', {"table:name": "List1"})
        # ET.SubElement(table, 'table:table-column')
        # ET.SubElement(table, 'table:table-row')
        return self.parse_element(root)

    def debug_cells(self) -> None:
        """Debug cells."""
        root = self.get_or_create_root()
        table = root.find(".//table:table", self.ns)
        row = ET.SubElement(table, 'table:table-row')
        cell = ET.SubElement(row, 'table:table-cell', {
            "office:value-type": "float",
            "office:value": "1",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, 'text:p')
        text.text = "1"
        cell = ET.SubElement(row, 'table:table-cell', {
            "office:value-type": "float",
            "office:value": "2",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, 'text:p')
        text.text = "2"
        cell = ET.SubElement(row, 'table:table-cell', {
            "office:value-type": "float",
            "office:value": "3",
            "calcext:value-type": "float",
        })
        text = ET.SubElement(cell, 'text:p')
        text.text = "3"

        cell = ET.SubElement(row, 'table:table-cell', {
            "office:value-type": "string",
            "calcext:value-type": "string",
        })
        text = ET.SubElement(cell, 'text:p')
        text.text = "Matěj"
