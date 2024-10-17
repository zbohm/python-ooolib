import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Styles(OpenDocumentMixin):

    filename = "styles.xml"

    def create(self) -> Element:
        """Create styles."""
        root = ET.Element(self.qname('office:document-styles'), {
            "office:version": self.version,
        })
        ET.SubElement(root, self.qname('office:styles'))
        return root
