from datetime import datetime
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin
from .version import VERSION


class Meta(OpenDocumentMixin):

    filename = "meta.xml"

    def create(self) -> Element:
        """Create meta."""
        root = self.create_element("office:document-meta", {"office:version": self.version})
        meta = self.create_sub_element(root, 'office:meta')
        self.create_sub_element(meta, "meta:creation-date", value=datetime.now().isoformat())
        self.create_sub_element(meta, "meta:generator", value=f'ooolib-python=={VERSION}')
        return root

    def get_or_create_root(self) -> Element:
        """Build meta."""
        if self.root is None:
            self.root = self.create()
        else:
            self.set_descendant_element_value("office:meta/dc:date", datetime.now().isoformat())
        return self.root
