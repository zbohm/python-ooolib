from xml.etree.ElementTree import Element

from .content import Content


class Write(Content):

    def create(self) -> Element:
        """Create content."""
        root = self.create_element("office:document-content", {"office:version": self.version})
        body = self.create_sub_element(root, "office:body")
        self.create_sub_element(body, "office:text")
        return root
