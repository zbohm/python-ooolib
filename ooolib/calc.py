import xml.etree.ElementTree as ET
from datetime import datetime
from io import BytesIO
from xml.etree.ElementTree import Element

VERSION = "1.0.0"


class LOCalc:
    """LibreOffice Calc."""

    ns = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }
    encoding = "utf-8"

    def __init__(self):
        self.meta: Element = None
        self.settings: Element = None
        self.styles: Element = None
        self.content: Element = None
        self.version = "1.2"

    def parse_element(self, element: Element) -> Element:
        """Parse element."""
        doc = ET.parse(BytesIO(ET.tostring(element, encoding=self.encoding)))
        return doc.getroot()

    def create_meta(self) -> Element:
        """Create meta."""
        root = ET.Element('office:document-meta', {
            "xmlns:office": self.ns["office"],
            "xmlns:meta": self.ns["meta"],
            "office:version": self.version,
        })
        meta = ET.SubElement(root, 'office:meta')
        creation_date = ET.SubElement(meta, 'meta:creation-date')
        creation_date.text = datetime.now().isoformat()
        ET.SubElement(meta, 'meta:generator', text=f'python-ooolib=={VERSION}')
        return self.parse_element(root)

    def build_meta(self) -> Element:
        """Build meta."""
        if self.meta is None:
            self.meta = self.create_meta()
        else:
            creation_date = self.meta.find("office:meta/meta:creation-date", self.ns)
            if creation_date is not None:
                creation_date.text = datetime.now().isoformat()
        return self.meta

    def create_settings(self) -> Element:
        """Create settings."""

    def create_styles(self) -> Element:
        """Create styles."""

    def create_content(self) -> Element:
        """Create content."""

    def load(self, filename: str) -> None:
        """Load document from filename."""
        if hasattr(ET, "register_namespace"):
            for name, uri in self.ns.items():
                ET.register_namespace(name, uri)
        doc = ET.parse(filename)
        self.meta = doc.getroot()

    def save(self, filename: str) -> None:
        """Save document into filename."""
        meta = self.build_meta()
        body = ET.tostring(meta, encoding='utf-8', xml_declaration=True)
        print(body.decode("utf-8"))


if __name__ == "__main__":
    calc = LOCalc()
    # calc.load('template/meta-f.xml')
    calc.save("test.ods")
