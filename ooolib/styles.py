from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Styles(OpenDocumentMixin):

    filename = "styles.xml"

    def create(self) -> Element:
        """Create styles."""
        root = self.create_element("office:document-styles", {"office:version": self.version})
        self.create_sub_element(root, "office:styles")
        return root
