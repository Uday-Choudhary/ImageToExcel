"""
Vision-based extraction strategy using Groq's Llama Vision API.

Handles image encoding, API communication, JSON parsing, and
optional file-based caching of results.
"""

from __future__ import annotations

import base64
import json
import logging
import os
from typing import Any, Optional

from core.constants import DEFAULT_TEMPERATURE, DEFAULT_VISION_MODEL, VISION_DATA_DIR, get_mime_type
from core.exceptions import ExtractionError, InvalidJSONError
from core.groq_client import create_groq_client
from core.prompts import VISION_EXTRACTION_PROMPT
from extractors.base import BaseExtractor

logger = logging.getLogger(__name__)


class VisionExtractor(BaseExtractor):
    """Extracts structured data from images using Groq's Llama Vision API.

    Attributes:
        model: The Llama Vision model identifier.
        temperature: Sampling temperature for the model.
        save_json: Whether to persist intermediate JSON to disk.
        api_key: Optional explicit API key (overrides env/secrets).
    """

    def __init__(
        self,
        model: str = DEFAULT_VISION_MODEL,
        temperature: float = DEFAULT_TEMPERATURE,
        save_json: bool = False,
        api_key: Optional[str] = None,
    ) -> None:
        self.model = model
        self.temperature = temperature
        self.save_json = save_json
        self._api_key = api_key

    @staticmethod
    def _encode_image_bytes(image_bytes: bytes) -> str:
        """Encode raw image bytes to a base64 string."""
        return base64.b64encode(image_bytes).decode("utf-8")

    @staticmethod
    def _encode_image_file(image_path: str) -> str:
        """Read and encode an image file to a base64 string."""
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    def extract_from_image(
        self,
        image_path: str = "",
        image_bytes: Optional[bytes] = None,
        filename: Optional[str] = None,
    ) -> Optional[dict[str, Any]]:
        """Extract structured data from an image via the Vision API.

        Args:
            image_path: Path to the image file (used if image_bytes not provided).
            image_bytes: Optional raw bytes of the image (for Streamlit uploads).
            filename: Original filename for MIME type detection and logging.

        Returns:
            Parsed JSON dict with extracted data, or None on failure.
        """
        fname = filename or os.path.basename(image_path)

        # Encode image
        if image_bytes:
            b64 = self._encode_image_bytes(image_bytes)
        elif image_path and os.path.exists(image_path):
            b64 = self._encode_image_file(image_path)
        else:
            logger.error("No image data provided for: %s", fname)
            return None

        mime = get_mime_type(fname)
        image_data_url = f"data:{mime};base64,{b64}"

        try:
            client = create_groq_client(self._api_key)

            completion = client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": VISION_EXTRACTION_PROMPT},
                            {"type": "image_url", "image_url": {"url": image_data_url}},
                        ],
                    }
                ],
                temperature=self.temperature,
                response_format={"type": "json_object"},
                stream=False,
            )

            content = completion.choices[0].message.content
            data = json.loads(content)

            # Optionally save to disk (for CLI pipeline)
            if self.save_json:
                self._persist_json(fname, data)

            logger.info("Successfully extracted data from: %s", fname)
            return data

        except json.JSONDecodeError:
            logger.error("Model returned invalid JSON for: %s", fname)
            return None
        except Exception as e:
            logger.error("API error for %s: %s", fname, e)
            return None

    def _persist_json(self, filename: str, data: dict) -> str:
        """Save extraction result to a JSON file.

        Args:
            filename: Original image filename.
            data: The parsed JSON response.

        Returns:
            Path to the saved JSON file.
        """
        os.makedirs(VISION_DATA_DIR, exist_ok=True)
        stem = filename.rsplit(".", 1)[0]
        json_path = os.path.join(VISION_DATA_DIR, f"{stem}_vision.json")

        with open(json_path, "w") as f:
            json.dump(data, f, indent=2)

        logger.debug("Saved vision data: %s", json_path)
        return json_path
