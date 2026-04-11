"""
Abstract base class for data extraction strategies.

Implements the Strategy pattern so that VisionExtractor and OCRExtractor
can be swapped transparently by the pipeline and Streamlit UI.
"""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from typing import Any, Optional

logger = logging.getLogger(__name__)


class BaseExtractor(ABC):
    """Abstract base class defining the extraction interface.

    Subclasses must implement `extract_from_image()` to provide
    their specific extraction logic (Vision API, OCR, etc.).
    """

    @abstractmethod
    def extract_from_image(
        self,
        image_path: str,
        image_bytes: Optional[bytes] = None,
        filename: Optional[str] = None,
    ) -> Optional[dict[str, Any]]:
        """Extract structured data from a single image.

        Args:
            image_path: Path to the image file on disk.
            image_bytes: Optional raw image bytes (used by Streamlit
                         where the file isn't saved to disk).
            filename: Original filename (used for MIME detection and logging).

        Returns:
            A dict containing the extracted data in the standard schema:
            {
                "document_summary": {...},
                "entities": {...},
                "tables": [{...}]
            }
            Returns None if extraction fails.
        """
        ...

    def extract_batch(
        self,
        image_paths: list[str],
        base_dir: str = "input",
    ) -> list[dict[str, Any]]:
        """Extract data from multiple images.

        Args:
            image_paths: List of image filenames (relative to base_dir).
            base_dir: Directory containing the images.

        Returns:
            List of successfully extracted result dicts.
        """
        import os

        results: list[dict[str, Any]] = []
        for img_name in image_paths:
            full_path = os.path.join(base_dir, img_name)
            logger.info("Processing: %s", img_name)
            data = self.extract_from_image(full_path, filename=img_name)
            if data:
                results.append(data)
            else:
                logger.warning("No data extracted from: %s", img_name)
        return results
