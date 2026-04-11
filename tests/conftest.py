"""Shared test fixtures for the ImageToExcel test suite."""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any

import pytest


@pytest.fixture
def sample_ocr_data() -> list[dict[str, Any]]:
    """Sample EasyOCR output data for testing spatial extraction."""
    return [
        {"bbox": [[50, 10], [150, 10], [150, 30], [50, 30]], "text": "Description", "confidence": 0.95},
        {"bbox": [[200, 10], [250, 10], [250, 30], [200, 30]], "text": "Qty", "confidence": 0.92},
        {"bbox": [[300, 10], [370, 10], [370, 30], [300, 30]], "text": "Price", "confidence": 0.90},
        {"bbox": [[420, 10], [500, 10], [500, 30], [420, 30]], "text": "Total", "confidence": 0.88},
        # Row 1
        {"bbox": [[50, 50], [180, 50], [180, 70], [50, 70]], "text": "Widget A", "confidence": 0.91},
        {"bbox": [[210, 50], [240, 50], [240, 70], [210, 70]], "text": "2", "confidence": 0.93},
        {"bbox": [[310, 50], [360, 50], [360, 70], [310, 70]], "text": "$10.00", "confidence": 0.89},
        {"bbox": [[430, 50], [490, 50], [490, 70], [430, 70]], "text": "$20.00", "confidence": 0.87},
        # Row 2
        {"bbox": [[50, 90], [180, 90], [180, 110], [50, 110]], "text": "Gadget B", "confidence": 0.90},
        {"bbox": [[210, 90], [240, 90], [240, 110], [210, 110]], "text": "5", "confidence": 0.94},
        {"bbox": [[310, 90], [360, 90], [360, 110], [310, 110]], "text": "$5.50", "confidence": 0.88},
        {"bbox": [[430, 90], [490, 90], [490, 110], [430, 110]], "text": "$27.50", "confidence": 0.86},
    ]


@pytest.fixture
def sample_ocr_json_path(sample_ocr_data: list[dict], tmp_path) -> str:
    """Write sample OCR data to a temporary JSON file and return the path."""
    json_path = tmp_path / "test_easyocr.json"
    with open(json_path, "w") as f:
        json.dump(sample_ocr_data, f)
    return str(json_path)


@pytest.fixture
def sample_vision_result() -> dict[str, Any]:
    """Sample Vision API extraction result."""
    return {
        "document_summary": {"style": "printed", "domain": "invoice"},
        "entities": {"company": "Acme Corp", "date": "2026-01-15"},
        "tables": [
            {
                "table_description": "Invoice Items",
                "headers": ["Description", "Qty", "Price", "Total"],
                "rows": [
                    {"Description": "Widget A", "Qty": "2", "Price": "$10.00", "Total": "$20.00"},
                    {"Description": "Gadget B", "Qty": "5", "Price": "$5.50", "Total": "$27.50"},
                ],
                "validation": {"math_check": "passed", "notes": "All totals match"},
            }
        ],
    }


@pytest.fixture
def sample_vision_results(sample_vision_result: dict) -> list[tuple[str, dict]]:
    """List of (sheet_name, data) tuples for Excel generation tests."""
    return [("TestInvoice", sample_vision_result)]
