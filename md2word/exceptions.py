class MD2WordError(Exception):
    """Base exception for MD2Word."""

class ParseError(MD2WordError):
    """Raised when markdown parsing fails."""

class RenderError(MD2WordError):
    """Raised when docx rendering fails."""

class ImageNotFoundError(MD2WordError):
    """Raised when an image file is not found."""
