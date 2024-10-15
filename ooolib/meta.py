import xml.etree.ElementTree as ET
from datetime import datetime
from xml.etree.ElementTree import Element

from .mixin import BaseMixin


class Meta(BaseMixin):

    filename = "meta.xml"

    def __init__(self):
        self.root: Element = None

    def create(self) -> Element:
        """Create meta."""
        root = ET.Element('office:document-meta', {
            "xmlns:office": self.ns["office"],
            "xmlns:meta": self.ns["meta"],
            "xmlns:dc": self.ns["dc"],
            "office:version": self.version,
        })
        meta = ET.SubElement(root, 'office:meta')
        creation_date = ET.SubElement(meta, 'meta:creation-date')
        creation_date.text = datetime.now().isoformat()
        version = self.get_version()
        ET.SubElement(meta, 'meta:generator', text=f'ooolib-python=={version}')
        return self.parse_element(root)

    def get_or_create_root(self) -> Element:
        """Build meta."""
        if self.root is None:
            self.root = self.create()
        else:
            self.set_value("office:meta/dc:date", datetime.now().isoformat())
            # self.set_value("office:meta/meta:editing-cycles", "2")
            # self.set_value("office:meta/meta:editing-duration", "PT38S")
        return self.root
