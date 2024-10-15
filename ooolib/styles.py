import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import BaseMixin


class Styles(BaseMixin):

    filename = "styles.xml"

    def __init__(self):
        self.root: Element = None

    def create(self) -> Element:
        """Create styles."""
        root = ET.Element('office:document-styles', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        ET.SubElement(root, 'office:styles')
        return self.parse_element(root)
