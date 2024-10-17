class OooLibException(Exception):
    """OooLib exception."""


class UnexpectedMimetype(OooLibException):
    """Unexpected mimetype."""


class ElementNotFound(OooLibException):
    """Element not found."""


class InvalidCellPosition(OooLibException):
    """Invalid cell position."""


class CellPositionOutOfRange(InvalidCellPosition):
    """Cell position is out of range."""
