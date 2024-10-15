import time
import xml.etree.ElementTree as ET
import zipfile

from .manifest import Manifest
from .meta import Meta
from .mixin import MainMixin
from .settings import Settings
from .spreadsheet import Spreadsheet
from .styles import Styles


class Calc(MainMixin):
    """LibreOffice Calc."""

    def __init__(self):
        self.manifest = Manifest()
        self.meta = Meta()
        self.settings = Settings()
        self.styles = Styles()
        self.sheet = Spreadsheet()

    def load(self, filename: str) -> None:
        """Load document from filename."""
        handle = zipfile.ZipFile(filename)
        try:
            self.meta.read(handle)
            self.manifest.read(handle)
            self.settings.read(handle)
            self.styles.read(handle)
            self.sheet.read(handle)
        finally:
            handle.close()

    def save(self, filename: str) -> None:
        """Save document into filename."""
        if hasattr(ET, "register_namespace"):
            for name, uri in self.meta.ns.items():
                ET.register_namespace(name, uri)

        self.sheet.debug_cells()  # DEBUG

        localtime = time.localtime()[:6]
        handle = zipfile.ZipFile(filename, "w")
        try:
            self.meta.write(handle, localtime)
            self.write_content(handle, localtime, "mimetype", self.sheet.mimetype)
            self.manifest.write(handle, localtime)
            self.settings.write(handle, localtime)
            self.styles.write(handle, localtime)
            self.sheet.write(handle, localtime)
        finally:
            handle.close()
