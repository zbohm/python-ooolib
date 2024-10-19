import re
import time
import xml.etree.ElementTree as ET
import zipfile

from .calc import Calc as CalcContent
from .content import Content, FileEntry
from .exceptions import UnexpectedMimetype
from .manifest import Manifest
from .meta import Meta
from .mixin import BaseMixin
from .settings import Settings
from .styles import Styles
from .write import Write as WriteContent


class OpenDocument(BaseMixin):
    """LibreOffice document."""

    mimetype: str

    def __init__(self):
        self.manifest = Manifest(self)
        self.meta = Meta(self)
        self.settings = Settings(self)
        self.styles = Styles(self)
        self.content = Content(self)
        self.payload: dict[str, FileEntry] = dict()

    def __getattr__(self, name):
        return getattr(self.content, name)

    def get_default_payload(self):
        """Get default payload."""
        return {
            self.manifest.filename: self.manifest,
            self.meta.filename: self.meta,
            self.settings.filename: self.settings,
            self.styles.filename: self.styles,
            self.content.filename: self.content,
        }

    def load(self, filename: str) -> None:
        """Load document from filename."""
        handle = zipfile.ZipFile(filename)
        payload = self.get_default_payload()
        try:
            mimetype = handle.read("mimetype").decode("utf-8")
            if self.mimetype is None:
                self.mimetype = mimetype
            else:
                if self.mimetype != mimetype:
                    raise UnexpectedMimetype(mimetype)
            manifest = payload.pop(self.manifest.filename)
            manifest.read(handle)
            self.payload[manifest.filename] = manifest
            for path, mimetype in manifest.get_file_entries():
                if re.search(r"\.\w+$", path):
                    file_entry = payload.get(path, FileEntry(self))
                    file_entry.filename = path
                    file_entry.mimetype = mimetype
                    file_entry.read(handle)
                    self.payload[path] = file_entry
        finally:
            handle.close()

    def save(self, filename: str) -> None:
        """Save document into filename."""
        for name, uri in self.meta.ns.items():
            ET.register_namespace(name, uri)
        localtime = time.localtime()[:6]
        handle = zipfile.ZipFile(filename, "w")
        if not self.payload:
            self.payload = self.get_default_payload()
        try:
            self.write_content(handle, localtime, "mimetype", self.mimetype.encode())
            for entry in self.payload.values():
                entry.write(handle, localtime)
        finally:
            handle.close()


class Calc(OpenDocument):
    """LibreOffice Calc."""

    mimetype = "application/vnd.oasis.opendocument.spreadsheet"

    def __init__(self):
        super().__init__()
        self.content = CalcContent(self)


class Write(OpenDocument):
    """LibreOffice Write."""

    mimetype = "application/vnd.oasis.opendocument.text"

    def __init__(self):
        super().__init__()
        self.content = WriteContent(self)
