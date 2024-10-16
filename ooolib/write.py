import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .content import Content


class Write(Content):

    def create(self) -> Element:
        """Create content."""
        root = ET.Element(self.qualify('office:document-content'), {
            "office:version": self.version,
        })
        body = ET.SubElement(root, self.qualify('office:body'))
        ET.SubElement(body, self.qualify('office:text'))
        return root
