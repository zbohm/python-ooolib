import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Settings(OpenDocumentMixin):

    filename = "settings.xml"

    def create(self) -> Element:
        """Create settings."""
        root = ET.Element(self.qname('office:document-settings'), {
            "office:version": self.version,
        })
        ET.SubElement(root, self.qname('office:settings'))
        return root
