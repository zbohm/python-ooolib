import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
from typing import Callable
from xml.etree.ElementTree import Element

localtimeType = tuple[int, int, int, int, int, int]  # year, month, day, hour, min, sec


class BaseMixin:

    def write_content(self, handle: zipfile.ZipFile, localtime: localtimeType, filename: str, content: bytes) -> None:
        """Write content."""
        info = zipfile.ZipInfo(filename)
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, content)


class RootMixin(BaseMixin):

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
        "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
        "dc": "http://purl.org/dc/elements/1.1/",
    }
    root: Element
    create: Callable
    filename: str

    def __init__(self, document: "ooolib.document.OpenDocument"):
        self.document = document
        self.root: Element = None

    def qualify(self, prefied_name: str) -> str:
        """Create qualified xml element name."""
        prefix, name = prefied_name.split(":")
        return f"{{{self.ns[prefix]}}}{name}"

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

    def set_value(self, xpath: str, value: str) -> None:
        """Set value to the xpath."""
        element = self.root.find(xpath, self.ns)
        if element is None:
            ancestor, name = xpath.rsplit('/', 1)
            parent = self.root.find(ancestor, self.ns)
            element = ET.SubElement(self.qualify(parent), self.qualify(name))
        element.text = value
