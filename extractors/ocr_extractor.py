"""
OCR-based extraction strategy using EasyOCR + spatial table analysis.

Handles image preprocessing, OCR execution, confidence comparison
between original and preprocessed versions, and spatial table
reconstruction via SpatialTableExtractor.
"""

from __future__ import annotations

import json
import logging
import os
from typing import Any, Optional

import easyocr
import numpy as np

from core.constants import MIN_OCR_CONFIDENCE, OCR_DATA_DIR, PREPROCESSED_DIR
from core.image_preprocessor import preprocess_image
from extractors.base import BaseExtractor
from extractors.spatial_table import SpatialTableExtractor

logger = logging.getLogger(__name__)


class NumpyEncoder(json.JSONEncoder):
    """JSON encoder that handles NumPy types gracefully."""

    def default(self, obj: Any) -> Any:
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super().default(obj)


class OCRExtractor(BaseExtractor):
    """Extracts structured data from images using EasyOCR + spatial analysis.

    Uses lazy-loaded EasyOCR reader for efficiency and compares
    OCR results between original and preprocessed images to pick
    the higher-confidence result.

    Attributes:
        gpu: Whether to use GPU acceleration for EasyOCR.
        languages: List of language codes for EasyOCR.
    """

    _reader: Optional[easyocr.Reader] = None  # Class-level lazy singleton

    def __init__(
        self,
        gpu: bool = False,
        languages: Optional[list[str]] = None,
    ) -> None:
        self.gpu = gpu
        self.languages = languages or ["en"]
        self._spatial_extractor = SpatialTableExtractor()

    def _get_reader(self) -> easyocr.Reader:
        """Get or create the lazy-loaded EasyOCR reader."""
        if OCRExtractor._reader is None:
            logger.info("Loading EasyOCR model...")
            OCRExtractor._reader = easyocr.Reader(self.languages, gpu=self.gpu)
        return OCRExtractor._reader

    def _run_easyocr(self, image_path: str) -> tuple[list[dict], float]:
        """Run EasyOCR with optimized parameters for document images.

        Args:
            image_path: Path to the image to process.

        Returns:
            Tuple of (formatted_results, average_confidence).
        """
        reader = self._get_reader()

        results = reader.readtext(
            image_path,
            detail=1,
            paragraph=False,
            min_size=10,
            text_threshold=0.6,
            low_text=0.3,
            width_ths=0.7,
            mag_ratio=1.5,
        )

        formatted: list[dict] = []
        total_conf = 0.0

        for bbox, text, conf in results:
            formatted.append({"bbox": bbox, "text": text, "confidence": conf})
            total_conf += conf

        avg_conf = total_conf / len(formatted) if formatted else 0.0
        return formatted, avg_conf

    def _run_ocr_dual(
        self, image_name: str, base_dir: str
    ) -> Optional[str]:
        """Run OCR on both original and preprocessed images, keep best.

        Args:
            image_name: Filename of the image.
            base_dir: Directory containing original images.

        Returns:
            Path to the saved JSON file, or None if no text detected.
        """
        original_path = os.path.join(base_dir, image_name)
        preprocessed_path = os.path.join(PREPROCESSED_DIR, image_name)

        # Ensure preprocessed version exists
        if not os.path.exists(preprocessed_path):
            preprocess_image(original_path, preprocessed_path)

        os.makedirs(OCR_DATA_DIR, exist_ok=True)
        best_results: list[dict] = []
        best_conf = 0.0
        best_source = ""

        # Run on original
        if os.path.exists(original_path):
            results, conf = self._run_easyocr(original_path)
            logger.debug("  Original: %d detections, conf=%.3f", len(results), conf)
            if conf > best_conf:
                best_results, best_conf, best_source = results, conf, "original"

        # Run on preprocessed
        if os.path.exists(preprocessed_path):
            results, conf = self._run_easyocr(preprocessed_path)
            logger.debug("  Preprocessed: %d detections, conf=%.3f", len(results), conf)
            if conf > best_conf:
                best_results, best_conf, best_source = results, conf, "preprocessed"

        if not best_results:
            logger.warning("No text detected in: %s", image_name)
            return None

        logger.info(
            "Using %s for %s (%d items, conf=%.3f)",
            best_source, image_name, len(best_results), best_conf,
        )

        # Save JSON
        stem = image_name.rsplit(".", 1)[0]
        json_path = os.path.join(OCR_DATA_DIR, f"{stem}_easyocr.json")
        with open(json_path, "w") as f:
            json.dump(best_results, f, cls=NumpyEncoder, indent=2)

        return json_path

    def extract_from_image(
        self,
        image_path: str = "",
        image_bytes: Optional[bytes] = None,
        filename: Optional[str] = None,
    ) -> Optional[dict[str, Any]]:
        """Extract structured data from an image using EasyOCR + spatial analysis.

        Args:
            image_path: Path to the image file.
            image_bytes: Unused (OCR requires file paths).
            filename: Optional filename override.

        Returns:
            Extraction result dict or None.
        """
        fname = filename or os.path.basename(image_path)
        base_dir = os.path.dirname(image_path) or "input"

        # Step 1: Run OCR
        json_path = self._run_ocr_dual(fname, base_dir)
        if not json_path:
            return None

        # Step 2: Spatial table extraction
        data = self._spatial_extractor.extract_full_data(json_path)
        if not data or not data.get("table", {}).get("rows"):
            logger.warning("No structured data found in: %s", fname)
            return None

        # Add sheet_name for Excel builder
        data["sheet_name"] = fname.rsplit(".", 1)[0]
        return data
