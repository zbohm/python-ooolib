import xml.etree.ElementTree as ET
from typing import Iterator, cast
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Manifest(OpenDocumentMixin):

    filename = "META-INF/manifest.xml"

    def create(self) -> Element:
        """Create manifest."""
        root = self.create_element("manifest:manifest", {"office:version": self.version})
        self.create_sub_element(root, 'manifest:file-entry', {
            "manifest:full-path": "/",
            "manifest:media-type": "application/vnd.oasis.opendocument.spreadsheet",
            "manifest:version": self.version,
        })
        for entry in self.document.payload.values():
            self.create_sub_element(root, "manifest:file-entry", {
                "manifest:full-path": entry.filename,
                "manifest:media-type": entry.mimetype,
            })
        return root

    def get_content(self) -> bytes:
        """Get content."""
        return ET.tostring(self.create(), encoding='utf-8', xml_declaration=True)

    def get_file_entries(self) -> Iterator[tuple[str, str]]:
        """Get file entries."""
        for entry in self.root.findall("manifest:file-entry", self.ns):
            yield (
                cast(str, entry.get(self.qname("manifest:full-path"))),
                cast(str, entry.get(self.qname("manifest:media-type"))),
            )
