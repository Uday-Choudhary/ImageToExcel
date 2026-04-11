"""
Custom exception classes for the ImageToExcel pipeline.

Using project-specific exceptions instead of generic ones
enables precise error handling and clearer logging.
"""

from __future__ import annotations


class ImageToExcelError(Exception):
    """Base exception for all ImageToExcel errors."""


class APIKeyMissingError(ImageToExcelError):
    """Raised when the Groq API key is not configured."""

    def __init__(self, message: str = "GROQ_API_KEY not found. Add it to Streamlit secrets or your .env file.") -> None:
        super().__init__(message)


class ExtractionError(ImageToExcelError):
    """Raised when data extraction fails for an image."""

    def __init__(self, filename: str, reason: str = "Unknown error") -> None:
        self.filename = filename
        self.reason = reason
        super().__init__(f"Extraction failed for '{filename}': {reason}")


class InvalidJSONError(ExtractionError):
    """Raised when the model returns invalid JSON."""

    def __init__(self, filename: str) -> None:
        super().__init__(filename, "Model returned invalid JSON")


class ImageNotFoundError(ImageToExcelError):
    """Raised when a specified image file does not exist."""

    def __init__(self, path: str) -> None:
        self.path = path
        super().__init__(f"Image not found: {path}")


class NoDataExtractedError(ImageToExcelError):
    """Raised when no structured data was extracted from any image."""
