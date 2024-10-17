from typing import Iterator
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import OpenDocumentMixin


class Manifest(OpenDocumentMixin):

    filename = "META-INF/manifest.xml"

    def create(self) -> Element:
        """Create manifest."""
        root = ET.Element(self.qualify('manifest:manifest'), {
            "manifest:version": self.version,
        })
        ET.SubElement(root, self.qualify("manifest:file-entry"), {
            "manifest:full-path": "/",
            "manifest:media-type": "application/vnd.oasis.opendocument.spreadsheet",
            "manifest:version": self.version,
        })

        for name in ("meta", "styles", "content", "settings"):
            ET.SubElement(root, self.qualify("manifest:file-entry"), {
                "manifest:full-path": f"{name}.xml",
                "manifest:media-type": "text/xml",
            })
        return root

    def get_file_entries(self) -> Iterator[tuple[str, str]]:
        """Get file entries."""
        for entry in self.root.findall("manifest:file-entry", self.ns):
            yield (
                entry.get(self.qualify("manifest:full-path")),
                entry.get(self.qualify("manifest:media-type")),
            )
