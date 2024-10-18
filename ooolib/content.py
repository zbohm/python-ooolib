from xml.etree.ElementTree import Element

from .mixin import FileEntryMixin, OpenDocumentMixin


class FileEntry(FileEntryMixin):
    """File Entry."""


class Content(OpenDocumentMixin):

    filename = "content.xml"

    def create(self) -> Element:
        """Create content."""
        root = self.create_element("office:document-content", {"office:version": self.version})
        self.create_sub_element(root, "office:body")
        return root
