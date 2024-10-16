import time
import xml.etree.ElementTree as ET
import zipfile

from .calc import Sheet
from .content import Content
from .exceptions import UnexpectedMimetype
from .manifest import Manifest
from .meta import Meta
from .mixin import BaseMixin
from .settings import Settings
from .styles import Styles
from .write import Write


class OpenDocument(BaseMixin):
    """LibreOffice document."""

    mimetype: str

    def __init__(self):
        self.manifest = Manifest(self)
        self.meta = Meta(self)
        self.settings = Settings(self)
        self.styles = Styles(self)
        self.content = Content(self)

    def __getattr__(self, name):
        return getattr(self.content, name)

    def load(self, filename: str) -> None:
        """Load document from filename."""
        handle = zipfile.ZipFile(filename)
        try:
            mimetype = handle.read("mimetype").decode("utf-8")
            if self.mimetype is None:
                self.mimetype = mimetype
            else:
                if self.mimetype != mimetype:
                    raise UnexpectedMimetype(mimetype)
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

    def __init__(self):
        super().__init__()
        self.content = Sheet(self)

    def save(self, filename: str) -> None:
        self.content.debug_cells()
        super().save(filename)


class Write(OpenDocument):
    """LibreOffice Write."""

    mimetype = "application/vnd.oasis.opendocument.text"

    def __init__(self):
        super().__init__()
        self.content = Write(self)
