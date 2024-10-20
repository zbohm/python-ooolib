import re
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
from typing import TYPE_CHECKING, Callable, Optional, Union, cast
from xml.etree.ElementTree import Element

from .exceptions import ElementNotFound

localtimeType = tuple[int, int, int, int, int, int]  # year, month, day, hour, min, sec
attrsType = dict[str, str]
valueType = Union[str, int, float]


if TYPE_CHECKING:
    from .document import OpenDocument


class BaseMixin:

    def write_content(self, handle: zipfile.ZipFile, localtime: localtimeType, filename: str, content: bytes) -> None:
        """Write content."""
        info = zipfile.ZipInfo(filename)
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, content)


class RootMixin:

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

    def __init__(self):
        super().__init__()
        self.root: Element = cast(Element, None)

    def qname(self, prefix_and_name: str) -> str:
        """Create qualified xml element name."""
        parts = prefix_and_name.split(":", 1)
        if len(parts) == 1:
            return parts[0]
        prefix, name = parts
        if prefix == "xmlns":
            return prefix_and_name
        return str(ET.QName(self.ns[prefix], name))

    def lname(self, qualified_name: str) -> str:
        """Get local name."""
        return re.sub(r"^\{.+?\}", "", qualified_name)

    def root_find(self, name: str) -> Element:
        """Find in root element."""
        node = self.root.find(name, self.ns)
        if node is None:
            raise ElementNotFound(name)
        return node

    def set_attrs(self, element: Element, attrs: attrsType) -> None:
        """Set element attributes."""
        for key, value in attrs.items():
            element.set(self.qname(key), value)

    def create_element(
            self,
            name: str,
            attrs: Optional[attrsType] = None,
            value: Optional[valueType] = None,
    ) -> Element:
        """Create element."""
        attributes = {} if attrs is None else {self.qname(key): value for key, value in attrs.items()}
        element = ET.Element(self.qname(name), attributes)
        if value is not None:
            element.text = str(value)
        return element

    def create_sub_element(
            self,
            parent: Element,
            name: str,
            attrs: Optional[attrsType] = None,
            value: Optional[valueType] = None,
    ) -> Element:
        """Create sub element."""
        attributes = {} if attrs is None else {self.qname(key): value for key, value in attrs.items()}
        element = ET.SubElement(parent, self.qname(name), attributes)
        if value is not None:
            element.text = str(value)
        return element

    def get_or_create_element(
            self,
            parent: Element,
            name: str,
            attrs: Optional[attrsType] = None,
            value: Optional[valueType] = None,
            alternate_name: Optional[str] = None,
    ) -> Element:
        """Get or create element."""
        element = parent.find(name, self.ns)
        if element is None:
            element = parent.find(alternate_name, self.ns)
        if element is None:
            return self.create_sub_element(parent, name, attrs, value)
        else:
            if attrs is not None:
                for key, val in attrs.items():
                    element.set(self.qname(key), val)
            if value is not None:
                element.text = str(value)
        return element

    def set_descendant_element_value(self, parent_and_name: str, value: str) -> None:
        """Set value to the element of 'prefix:parent/prefix:name'."""
        element = self.root.find(parent_and_name, self.ns)
        if element is None:
            ancestor, name = parent_and_name.rsplit('/', 1)
            parent = self.root_find(ancestor)
            element = self.create_sub_element(parent, name)
        element.text = value


class FileEntryMixin(BaseMixin):

    filename: str
    mimetype = "text/xml"

    def __init__(self, document: "OpenDocument") -> None:
        super().__init__()
        self.document = document
        self.content: bytes = b""

    def read(self, handle: zipfile.ZipFile) -> None:
        """Read content from handle."""
        self.content = handle.read(self.filename)

    def get_content(self) -> bytes:
        """Get content."""
        return self.content

    def write(self, handle: zipfile.ZipFile, localtime: localtimeType) -> None:
        """Write content into the handle."""
        self.write_content(handle, localtime, self.filename, self.get_content())


class OpenDocumentMixin(FileEntryMixin, RootMixin):

    create: Callable

    def get_or_create_root(self) -> Element:
        """Get or create section."""
        if self.root is None:
            self.root = self.create()
        return self.root

    def read(self, handle: zipfile.ZipFile) -> None:
        """Read content from handle."""
        doc = ET.parse(BytesIO(handle.read(self.filename)))
        self.root = doc.getroot()

    def get_content(self) -> bytes:
        """Get content."""
        return ET.tostring(self.get_or_create_root(), encoding='utf-8', xml_declaration=True)
