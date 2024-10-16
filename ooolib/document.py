import time
import xml.etree.ElementTree as ET
import zipfile

from .content import Content
from .manifest import Manifest
from .meta import Meta
from .mixin import MainMixin
from .settings import Settings
from .styles import Styles


class OpenDocument(MainMixin):
    """LibreOffice document."""

    mimetype: str

    def __init__(self):
        self.manifest = Manifest()
        self.meta = Meta()
        self.settings = Settings()
        self.styles = Styles()
        self.content = Content()

    def load(self, filename: str) -> None:
        """Load document from filename."""
        handle = zipfile.ZipFile(filename)
        try:
            self.meta.read(handle)
            self.manifest.read(handle)
            self.settings.read(handle)
            self.styles.read(handle)
            self.content.read(handle)
        finally:
            handle.close()

    def save(self, filename: str) -> None:
        """Save document into filename."""
        if hasattr(ET, "register_namespace"):
            for name, uri in self.meta.ns.items():
                ET.register_namespace(name, uri)

        self.content.debug_cells()  # DEBUG

        localtime = time.localtime()[:6]
        handle = zipfile.ZipFile(filename, "w")
        try:
            self.meta.write(handle, localtime)
            self.write_content(handle, localtime, "mimetype", self.mimetype.encode())
            self.manifest.write(handle, localtime)
            self.settings.write(handle, localtime)
            self.styles.write(handle, localtime)
            self.content.write(handle, localtime)
        finally:
            handle.close()


class Calc(OpenDocument):
    """LibreOffice Calc."""

    mimetype = "application/vnd.oasis.opendocument.spreadsheet"


class Write(OpenDocument):
    """LibreOffice Write."""

    mimetype = "application/vnd.oasis.opendocument.text"
