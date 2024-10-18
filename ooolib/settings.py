from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Settings(OpenDocumentMixin):

    filename = "settings.xml"

    def create(self) -> Element:
        """Create settings."""
        root = self.create_element("office:document-settings", {"office:version": self.version})
        self.create_sub_element(root, "office:settings")
        return root
