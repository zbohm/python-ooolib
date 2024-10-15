import xml.etree.ElementTree as ET
from datetime import datetime
from xml.etree.ElementTree import Element

from .mixin import BaseMixin

VERSION = "1.0.0"


class Meta(BaseMixin):

    filename = "meta.xml"

    def __init__(self):
        self.root: Element = None

    def create(self) -> Element:
        """Create meta."""
        root = ET.Element('office:document-meta', {
            "xmlns:office": self.ns["office"],
            "xmlns:meta": self.ns["meta"],
            "office:version": self.version,
        })
        meta = ET.SubElement(root, 'office:meta')
        creation_date = ET.SubElement(meta, 'meta:creation-date')
        creation_date.text = datetime.now().isoformat()
        ET.SubElement(meta, 'meta:generator', text=f'ooolib-python=={VERSION}')
        return self.parse_element(root)

    def get_or_create_root(self) -> Element:
        """Build meta."""
        if self.root is None:
            self.root = self.create()
        else:
            creation_date = self.root.find("office:meta/meta:creation-date", self.ns)
            if creation_date is not None:
                creation_date.text = datetime.now().isoformat()
        return self.root
