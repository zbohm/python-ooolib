class OooLibException(Exception):
    """OooLib exception."""


class UnexpectedMimetype(OooLibException):
    """Unexpected mimetype."""


class ElementNotFound(OooLibException):
    """Element not found."""
