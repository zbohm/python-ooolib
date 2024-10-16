import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import RootMixin


class Settings(RootMixin):

    filename = "settings.xml"

    def create(self) -> Element:
        """Create settings."""
        root = ET.Element('office:document-settings', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        ET.SubElement(root, 'office:settings')
        return self.parse_element(root)
