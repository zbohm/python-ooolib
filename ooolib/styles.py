import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import RootMixin


class Styles(RootMixin):

    filename = "styles.xml"

    def create(self) -> Element:
        """Create styles."""
        root = ET.Element(self.qualify('office:document-styles'), {
            "office:version": self.version,
        })
        ET.SubElement(root, self.qualify('office:styles'))
        return root
