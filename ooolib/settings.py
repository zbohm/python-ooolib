import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Settings(OpenDocumentMixin):

    filename = "settings.xml"

    def create(self) -> Element:
        """Create settings."""
        root = ET.Element(self.qualify('office:document-settings'), {
            "office:version": self.version,
        })
        ET.SubElement(root, self.qualify('office:settings'))
        return root
