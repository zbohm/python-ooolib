import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .mixin import BaseMixin


class Manifest(BaseMixin):

    filename = "META-INF/manifest.xml"

    def __init__(self):
        self.root: Element = None

    def create(self) -> Element:
        """Create manifest."""
        root = ET.Element('manifest:manifest', {
            "xmlns:manifest": self.ns["manifest"],
            "manifest:version": self.version,
        })
        ET.SubElement(root, "manifest:file-entry", {
            "manifest:full-path": "/",
            "manifest:media-type": "application/vnd.oasis.opendocument.spreadsheet",
            "manifest:version": self.version,
        })

        for name in ("meta", "styles", "content", "settings"):
            ET.SubElement(root, "manifest:file-entry", {
                "manifest:full-path": f"{name}.xml",
                "manifest:media-type": "text/xml",
            })
        return self.parse_element(root)
