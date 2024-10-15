import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from typing import Callable
from xml.etree.ElementTree import Element

localtimeType = tuple[int, int, int, int, int, int]


class MainMixin:

    def write_content(self, handle: zipfile.ZipFile, localtime: localtimeType, filename: str, content: bytes) -> None:
        info = zipfile.ZipInfo(filename)
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, content)


class BaseMixin(MainMixin):

    version = "1.2"
    encoding = "utf-8"
    ns = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
        "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    }
    root: Element
    create: Callable
    filename: str

    def parse_element(self, element: Element) -> Element:
        """Parse element."""
        doc = ET.parse(BytesIO(ET.tostring(element, encoding=self.encoding)))
        return doc.getroot()

    def get_or_create_root(self) -> Element:
        """Get or create section."""
        if self.root is None:
            self.root = self.create()
        return self.root

    def read(self, handle: zipfile.ZipFile) -> None:
        """Read content from handle."""
        doc = ET.parse(BytesIO(handle.read(self.filename)))
        self.root = doc.getroot()

    def write(self, handle: zipfile.ZipFile, localtime: localtimeType) -> None:
        """Write content into the handle."""
        self.write_content(
            handle,
            localtime,
            self.filename,
            ET.tostring(self.get_or_create_root(), encoding='utf-8', xml_declaration=True)
        )
