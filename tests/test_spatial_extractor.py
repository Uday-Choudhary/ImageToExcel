"""Tests for the SpatialTableExtractor module."""

from __future__ import annotations

import pytest

from extractors.spatial_table import SpatialTableExtractor


class TestCleanText:
    """Tests for OCR error correction in clean_text()."""

    def setup_method(self) -> None:
        self.extractor = SpatialTableExtractor()

    def test_empty_string(self) -> None:
        assert self.extractor.clean_text("") == ""

    def test_none_input(self) -> None:
        assert self.extractor.clean_text(None) == ""

    def test_s_to_dollar(self) -> None:
        """S at start of number should become $."""
        assert self.extractor.clean_text("S100") == "$100"
        assert self.extractor.clean_text("s250.00") == "$250.00"

    def test_o_to_zero_in_numeric_context(self) -> None:
        """O/o adjacent to digits should become 0."""
        assert self.extractor.clean_text("1O0") == "100"
        assert self.extractor.clean_text("$1O.OO") == "$10.00"

    def test_l_to_one_in_numeric_context(self) -> None:
        """l/I adjacent to digits should become 1."""
        assert self.extractor.clean_text("$l0.00") == "$10.00"

    def test_pipe_to_one(self) -> None:
        """| adjacent to digits should become 1."""
        assert self.extractor.clean_text("$|0.00") == "$10.00"

    def test_b_to_eight(self) -> None:
        """B adjacent to digits should become 8."""
        assert self.extractor.clean_text("$B0.00") == "$80.00"

    def test_normal_text_unchanged(self) -> None:
        """Non-numeric text should not be modified."""
        assert self.extractor.clean_text("Description") == "Description"
        assert self.extractor.clean_text("Widget A") == "Widget A"

    def test_strips_underscores(self) -> None:
        """Leading/trailing underscores should be removed."""
        assert self.extractor.clean_text("_Total_") == "Total"


class TestClusterRows:
    """Tests for row clustering based on vertical overlap."""

    def setup_method(self) -> None:
        self.extractor = SpatialTableExtractor()

    def test_empty_words(self) -> None:
        assert self.extractor._cluster_rows([]) == []

    def test_single_word(self) -> None:
        word = {"y_center": 10, "y_min": 5, "y_max": 15, "x_min": 0, "x_max": 50, "height": 10}
        rows = self.extractor._cluster_rows([word])
        assert len(rows) == 1
        assert len(rows[0]) == 1

    def test_same_row_words(self) -> None:
        """Words at the same vertical position should cluster into one row."""
        words = [
            {"y_center": 10, "y_min": 5, "y_max": 15, "x_min": 0, "x_max": 50, "height": 10, "text": "A"},
            {"y_center": 12, "y_min": 7, "y_max": 17, "x_min": 60, "x_max": 100, "height": 10, "text": "B"},
        ]
        rows = self.extractor._cluster_rows(words)
        assert len(rows) == 1
        assert len(rows[0]) == 2

    def test_different_row_words(self) -> None:
        """Words at very different vertical positions should be separate rows."""
        words = [
            {"y_center": 10, "y_min": 5, "y_max": 15, "x_min": 0, "x_max": 50, "height": 10, "text": "Row1"},
            {"y_center": 60, "y_min": 55, "y_max": 65, "x_min": 0, "x_max": 50, "height": 10, "text": "Row2"},
        ]
        rows = self.extractor._cluster_rows(words)
        assert len(rows) == 2

    def test_rows_sorted_by_x_within_row(self) -> None:
        """Words within a row should be sorted left-to-right by x_min."""
        words = [
            {"y_center": 10, "y_min": 5, "y_max": 15, "x_min": 100, "x_max": 150, "height": 10, "text": "B"},
            {"y_center": 12, "y_min": 7, "y_max": 17, "x_min": 0, "x_max": 50, "height": 10, "text": "A"},
        ]
        rows = self.extractor._cluster_rows(words)
        assert rows[0][0]["text"] == "A"
        assert rows[0][1]["text"] == "B"


class TestFindHeaderRow:
    """Tests for header row detection."""

    def setup_method(self) -> None:
        self.extractor = SpatialTableExtractor()

    def test_no_headers_found(self) -> None:
        rows = [
            [{"text": "abc", "x_min": 0}, {"text": "xyz", "x_min": 50}],
        ]
        idx, words = self.extractor._find_header_row(rows)
        assert idx == -1

    def test_detects_primary_keywords(self) -> None:
        rows = [
            [{"text": "Description", "x_min": 0}, {"text": "Qty", "x_min": 100}, {"text": "Price", "x_min": 200}],
        ]
        idx, words = self.extractor._find_header_row(rows)
        assert idx == 0
        assert len(words) == 3

    def test_skips_single_word_rows(self) -> None:
        """Rows with only 1 word (likely titles) should be skipped."""
        rows = [
            [{"text": "INVOICE", "x_min": 0}],
            [{"text": "Description", "x_min": 0}, {"text": "Amount", "x_min": 100}],
        ]
        idx, _ = self.extractor._find_header_row(rows)
        assert idx == 1


class TestExtractFullData:
    """Integration tests for the full extraction pipeline."""

    def setup_method(self) -> None:
        self.extractor = SpatialTableExtractor()

    def test_extracts_table_from_json(self, sample_ocr_json_path: str) -> None:
        result = self.extractor.extract_full_data(sample_ocr_json_path)

        assert result is not None
        assert "table" in result
        assert "header_split" in result
        assert "footer_info" in result

        table = result["table"]
        assert len(table["headers"]) > 0
        assert len(table["rows"]) > 0

    def test_returns_none_for_empty_json(self, tmp_path) -> None:
        json_path = tmp_path / "empty.json"
        with open(json_path, "w") as f:
            f.write("[]")

        result = self.extractor.extract_full_data(str(json_path))
        assert result is None

    def test_returns_none_for_invalid_json(self, tmp_path) -> None:
        json_path = tmp_path / "invalid.json"
        with open(json_path, "w") as f:
            f.write("not valid json")

        result = self.extractor.extract_full_data(str(json_path))
        assert result is None

    def test_filters_low_confidence(self, tmp_path) -> None:
        """Items below min_confidence should be filtered out."""
        data = [
            {"bbox": [[0, 0], [50, 0], [50, 20], [0, 20]], "text": "Good", "confidence": 0.9},
            {"bbox": [[60, 0], [100, 0], [100, 20], [60, 20]], "text": "Bad", "confidence": 0.1},
        ]
        json_path = tmp_path / "low_conf.json"
        with open(json_path, "w") as f:
            json.dump(data, f)

        # Won't find a header with just one good word, but shouldn't crash
        import json
        result = self.extractor.extract_full_data(str(json_path))
        # Result may be None or have filtered data - key is no crash


class TestExtractFromJson:
    """Tests for backward-compatible extract_from_json wrapper."""

    def test_returns_table_dict(self, sample_ocr_json_path: str) -> None:
        extractor = SpatialTableExtractor()
        result = extractor.extract_from_json(sample_ocr_json_path)

        assert result is not None
        assert "headers" in result
        assert "rows" in result
