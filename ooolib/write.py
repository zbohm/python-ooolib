import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .content import Content


class Write(Content):

    def create(self) -> Element:
        """Create content."""
        root = ET.Element(self.qname('office:document-content'), {
            "office:version": self.version,
        })
        body = ET.SubElement(root, self.qname('office:body'))
        ET.SubElement(body, self.qname('office:text'))
        return root
