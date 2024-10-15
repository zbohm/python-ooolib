import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from io import BytesIO
from xml.etree.ElementTree import Element

VERSION = "1.0.0"


class Calc:
    """LibreOffice Calc."""

    ns = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
        "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    }
    encoding = "utf-8"

    def __init__(self):
        self.manifest: Element = None
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
        ET.SubElement(meta, 'meta:generator', text=f'ooolib-python=={VERSION}')
        return self.parse_element(root)

    @property
    def section_meta(self) -> Element:
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
        root = ET.Element('office:document-settings', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        ET.SubElement(root, 'office:settings')
        return self.parse_element(root)

    @property
    def section_settings(self) -> Element:
        """Build settings."""
        if self.settings is None:
            self.settings = self.create_settings()
        return self.settings

    def create_styles(self) -> Element:
        """Create styles."""
        root = ET.Element('office:document-styles', {
            "xmlns:office": self.ns["office"],
            "office:version": self.version,
        })
        ET.SubElement(root, 'office:styles')
        return self.parse_element(root)

    @property
    def section_styles(self) -> Element:
        """Build styles."""
        if self.styles is None:
            self.styles = self.create_styles()
        return self.styles

    def create_content(self) -> Element:
        """Create content."""
        root = ET.Element('office:document-content', {
            "xmlns:office": self.ns["office"],
            "xmlns:table": self.ns["table"],
            "xmlns:style": self.ns["style"],
            "xmlns:fo": self.ns["fo"],
            "office:version": self.version,
        })
        automatic_styles = ET.SubElement(root, 'office:automatic-styles')
        # Column
        style = ET.SubElement(automatic_styles, 'style:style', {
            "style:name": "co1",
            "style:family": "table-column",
        })
        ET.SubElement(style, 'style:table-column-properties', {
             "style:column-width": "2.258cm",
            "fo:break-before": "auto",
        })
        # Row
        style = ET.SubElement(automatic_styles, 'style:style', {
            "style:name": "ro1",
            "style:family": "table-row",
        })
        ET.SubElement(style, 'style:table-row-properties', {
            "style:row-height": "0.452cm" ,
            "style:use-optimal-row-height": "true",
            "fo:break-before": "auto",
        })
        # Table
        style = ET.SubElement(automatic_styles, 'style:style', {
            "style:name": "ta1",
            "style:family": "table",
            "style:master-page-name": "Default",
        })
        ET.SubElement(style, 'style:table-properties', {
            "table:display": "true",
            "style:writing-mode": "lr-tb",
        })

        body = ET.SubElement(root, 'office:body')
        sheet = ET.SubElement(body, 'office:spreadsheet')
        table = ET.SubElement(sheet, 'table:table', {"table:name": "List1", "table:style-name": "ta1"})
        ET.SubElement(table, 'table:table-column', {"table:style-name": "co1", "table:default-cell-style": "Default"})
        row = ET.SubElement(table, 'table:table-row', {"table:style-name": "ro1"})
        ET.SubElement(row, 'table:table-cell')
        return self.parse_element(root)

    @property
    def section_content(self) -> Element:
        """Build content."""
        if self.content is None:
            self.content = self.create_content()
        return self.content

    def create_manifest(self) -> Element:
        """Create content."""
        ns = {
            "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
        }
        root = ET.Element('manifest:manifest', {
            "xmlns:manifest": ns["manifest"],
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

    @property
    def section_manifest(self):
        """Build manifest."""
        if self.manifest is None:
            self.manifest = self.create_manifest()
        return self.manifest

    def load(self, filename: str) -> None:
        """Load document from filename."""
        doc = ET.parse(filename)
        self.meta = doc.getroot()

    def save(self, filename: str) -> None:
        """Save document into filename."""
        if hasattr(ET, "register_namespace"):
            for name, uri in self.ns.items():
                ET.register_namespace(name, uri)

        # body = ET.tostring(self.section_meta, encoding='utf-8', xml_declaration=True)
        # print(body.decode("utf-8"))
        # body = ET.tostring(self.section_settings, encoding='utf-8', xml_declaration=True)
        # print(body.decode("utf-8"))
        # body = ET.tostring(self.section_styles, encoding='utf-8', xml_declaration=True)
        # print(body.decode("utf-8"))
        # body = ET.tostring(self.section_content, encoding='utf-8', xml_declaration=True)
        # print(body.decode("utf-8"))
        # body = ET.tostring(self.section_manifest, encoding='utf-8', xml_declaration=True)
        # print(body.decode("utf-8"))

        localtime = time.localtime()[:6]
        handle = zipfile.ZipFile(filename, "w")

        info = zipfile.ZipInfo("meta.xml")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, ET.tostring(self.section_meta, encoding='utf-8', xml_declaration=True))

        info = zipfile.ZipInfo("mimetype")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, "application/vnd.oasis.opendocument.text")

        info = zipfile.ZipInfo("META-INF/manifest.xml")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, ET.tostring(self.section_manifest, encoding='utf-8', xml_declaration=True))

        info = zipfile.ZipInfo("settings.xml")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, ET.tostring(self.section_settings, encoding='utf-8', xml_declaration=True))

        info = zipfile.ZipInfo("styles.xml")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, ET.tostring(self.section_styles, encoding='utf-8', xml_declaration=True))

        info = zipfile.ZipInfo("content.xml")
        info.date_time = localtime
        info.compress_type = zipfile.ZIP_DEFLATED
        handle.writestr(info, ET.tostring(self.section_content, encoding='utf-8', xml_declaration=True))

        handle.close()


if __name__ == "__main__":
    calc = Calc()
    # calc.load('ooolib/template/meta-f.xml')
    calc.save("test.ods")
