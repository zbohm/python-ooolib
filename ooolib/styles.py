import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import RootMixin


class Styles(RootMixin):

    filename = "styles.xml"

    def create(self) -> Element:
        """Create styles."""
        root = ET.Element('office:document-styles', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        ET.SubElement(root, 'office:styles')
        return self.parse_element(root)
