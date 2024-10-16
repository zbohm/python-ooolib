import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Content(OpenDocumentMixin):

    filename = "content.xml"

    def create(self) -> Element:
        """Create content."""
        root = ET.Element(self.qualify('office:document-content'), {
            "office:version": self.version,
        })
        ET.SubElement(root, self.qualify('office:body'))
        return root
