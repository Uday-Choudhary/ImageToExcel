"""
Centralized constants for the ImageToExcel pipeline.

All directory paths, file extensions, model names, and tuning thresholds
live here to avoid scattering magic strings across modules.
"""

from __future__ import annotations

import os
from pathlib import Path

# ── Directory Paths ──────────────────────────────────────────────────────────
PROJECT_ROOT: Path = Path(__file__).resolve().parent.parent

INPUT_DIR: str = "input"
OUTPUT_DIR: str = "output"
PREPROCESSED_DIR: str = "preprocessed"
OCR_DATA_DIR: str = "ocr_data"
VISION_DATA_DIR: str = "vision_data"

# ── Supported Image Types ───────────────────────────────────────────────────
VALID_IMAGE_EXTENSIONS: tuple[str, ...] = (".jpg", ".jpeg", ".png")

# ── Groq / LLM Configuration ────────────────────────────────────────────────
DEFAULT_VISION_MODEL: str = "meta-llama/llama-4-scout-17b-16e-instruct"
DEFAULT_TEMPERATURE: float = 0.1

# ── Excel ────────────────────────────────────────────────────────────────────
MAX_SHEET_NAME_LENGTH: int = 31  # Excel specification limit
MAX_COLUMN_WIDTH: int = 50
HEADER_FILL_COLOR: str = "4F81BD"
SECONDARY_FILL_COLOR: str = "6F819D"

# ── OCR / Spatial Thresholds ────────────────────────────────────────────────
MIN_OCR_CONFIDENCE: float = 0.3
VERTICAL_OVERLAP_THRESHOLD: float = 0.35
DESKEW_MIN_ANGLE: float = 0.5
DESKEW_MAX_ANGLE: float = 15.0
MATH_VALIDATION_TOLERANCE: float = 0.05

# ── Validation Colors (Excel) ──────────────────────────────────────────────
VALIDATION_PASS_COLOR: str = "E2EFDA"
VALIDATION_FAIL_COLOR: str = "FFC7CE"

# ── MIME Types ──────────────────────────────────────────────────────────────
MIME_MAP: dict[str, str] = {
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "png": "image/png",
}


def get_mime_type(filename: str) -> str:
    """Return the MIME type for a given filename based on its extension.

    Args:
        filename: The filename or path to determine MIME type for.

    Returns:
        The MIME type string, defaulting to 'image/jpeg' for unknown types.
    """
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    return MIME_MAP.get(ext, "image/jpeg")
