"""Tests for the Groq client factory and supporting modules."""

from __future__ import annotations

import os
from unittest.mock import MagicMock, patch

import pytest

from core.constants import get_mime_type
from core.exceptions import (
    APIKeyMissingError,
    ExtractionError,
    ImageNotFoundError,
    InvalidJSONError,
)
from core.groq_client import _resolve_api_key, create_groq_client


class TestMimeType:
    """Tests for MIME type detection."""

    def test_jpg(self) -> None:
        assert get_mime_type("photo.jpg") == "image/jpeg"

    def test_jpeg(self) -> None:
        assert get_mime_type("photo.jpeg") == "image/jpeg"

    def test_png(self) -> None:
        assert get_mime_type("photo.png") == "image/png"

    def test_uppercase_extension(self) -> None:
        assert get_mime_type("PHOTO.PNG") == "image/png"

    def test_unknown_defaults_to_jpeg(self) -> None:
        assert get_mime_type("file.bmp") == "image/jpeg"

    def test_no_extension(self) -> None:
        assert get_mime_type("noext") == "image/jpeg"


class TestCustomExceptions:
    """Tests for custom exception classes."""

    def test_api_key_missing_default_message(self) -> None:
        with pytest.raises(APIKeyMissingError, match="GROQ_API_KEY not found"):
            raise APIKeyMissingError()

    def test_extraction_error(self) -> None:
        err = ExtractionError("test.jpg", "timeout")
        assert err.filename == "test.jpg"
        assert err.reason == "timeout"
        assert "test.jpg" in str(err)

    def test_invalid_json_error(self) -> None:
        err = InvalidJSONError("photo.png")
        assert err.filename == "photo.png"
        assert "invalid JSON" in str(err)

    def test_image_not_found_error(self) -> None:
        err = ImageNotFoundError("/path/to/img.jpg")
        assert err.path == "/path/to/img.jpg"


class TestResolveApiKey:
    """Tests for API key resolution logic."""

    def test_explicit_key_takes_priority(self) -> None:
        key = _resolve_api_key("gsk_test_key")
        assert key == "gsk_test_key"

    def test_strips_whitespace(self) -> None:
        key = _resolve_api_key("  gsk_test_key  ")
        assert key == "gsk_test_key"

    @patch.dict(os.environ, {"GROQ_API_KEY": "gsk_from_env"})
    def test_falls_back_to_env(self) -> None:
        key = _resolve_api_key()
        assert key == "gsk_from_env"

    @patch.dict(os.environ, {}, clear=True)
    def test_raises_when_no_key(self) -> None:
        # Remove env var and ensure no streamlit secrets
        with pytest.raises(APIKeyMissingError):
            _resolve_api_key()

    @patch.dict(os.environ, {"GROQ_API_KEY": "gsk_from_env"})
    def test_create_client_returns_groq_instance(self) -> None:
        client = create_groq_client()
        assert client is not None


class TestVisionExtractor:
    """Tests for VisionExtractor with mocked API calls."""

    @patch("extractors.vision_extractor.create_groq_client")
    def test_extract_from_bytes(self, mock_create_client) -> None:
        from extractors.vision_extractor import VisionExtractor

        # Mock the Groq client response
        mock_client = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [
            MagicMock(message=MagicMock(content='{"document_summary": {}, "entities": {}, "tables": []}'))
        ]
        mock_client.chat.completions.create.return_value = mock_completion
        mock_create_client.return_value = mock_client

        extractor = VisionExtractor(api_key="gsk_test")
        result = extractor.extract_from_image(
            image_bytes=b"\x89PNG\r\n\x1a\n",
            filename="test.png",
        )

        assert result is not None
        assert "document_summary" in result
        mock_client.chat.completions.create.assert_called_once()

    @patch("extractors.vision_extractor.create_groq_client")
    def test_returns_none_on_invalid_json(self, mock_create_client) -> None:
        from extractors.vision_extractor import VisionExtractor

        mock_client = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [MagicMock(message=MagicMock(content="not json"))]
        mock_client.chat.completions.create.return_value = mock_completion
        mock_create_client.return_value = mock_client

        extractor = VisionExtractor(api_key="gsk_test")
        result = extractor.extract_from_image(
            image_bytes=b"\x89PNG",
            filename="test.png",
        )
        assert result is None

    def test_returns_none_when_no_image_data(self) -> None:
        from extractors.vision_extractor import VisionExtractor

        extractor = VisionExtractor(api_key="gsk_test")
        result = extractor.extract_from_image(
            image_path="/nonexistent/path.jpg",
            filename="path.jpg",
        )
        assert result is None
