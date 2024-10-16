import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .content import Content


class Write(Content):

    def create(self) -> Element:
        """Create content."""
        root = ET.Element('office:document-content', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        body = ET.SubElement(root, 'office:body')
        ET.SubElement(body, 'office:text')
        return self.parse_element(root)
