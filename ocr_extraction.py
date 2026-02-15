"""
OCR Extraction
Runs EasyOCR to detect text and tables in images.
"""

import easyocr
import os
import json
import numpy as np


class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        return super().default(obj)

PREPROCESSED_DIR = "preprocessed"
OCR_DATA_DIR = "ocr_data"

# Lazy-loaded EasyOCR reader
_reader = None


def _get_reader():
    global _reader
    if _reader is None:
        print("  Loading EasyOCR model...")
        _reader = easyocr.Reader(['en'], gpu=False)
    return _reader


def _run_easyocr(image_path, reader):
    """Run EasyOCR with optimized params for table/document images."""
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

    formatted = []
    total_conf = 0.0
    for bbox, text, conf in results:
        formatted.append({"bbox": bbox, "text": text, "confidence": conf})
        total_conf += conf

    avg_conf = total_conf / len(formatted) if formatted else 0
    return formatted, avg_conf


def extract_text(image_name, base_dir="input"):
    """
    Extract text from an image using EasyOCR.
    Runs on both original and preprocessed versions, picks the best.
    Saves JSON to ocr_data/ folder.
    """
    reader = _get_reader()

    original_path = os.path.join(base_dir, image_name)
    preprocessed_path = os.path.join(PREPROCESSED_DIR, image_name)
    out_dir = OCR_DATA_DIR
    os.makedirs(out_dir, exist_ok=True)

    best_results, best_conf, best_source = [], 0.0, ""

    # Run on original
    if os.path.exists(original_path):
        results, conf = _run_easyocr(original_path, reader)
        print(f"    Original:      {len(results):3d} detections, conf={conf:.3f}")
        if conf > best_conf:
            best_results, best_conf, best_source = results, conf, "original"

    # Run on preprocessed
    if os.path.exists(preprocessed_path):
        results, conf = _run_easyocr(preprocessed_path, reader)
        print(f"    Preprocessed:  {len(results):3d} detections, conf={conf:.3f}")
        if conf > best_conf:
            best_results, best_conf, best_source = results, conf, "preprocessed"

    if not best_results:
        print(f"    ERROR: No text detected in {image_name}")
        return None

    print(f"    → Using {best_source} ({len(best_results)} items, conf={best_conf:.3f})")

    # Save JSON
    stem = image_name.rsplit('.', 1)[0]
    json_path = os.path.join(out_dir, f"{stem}_easyocr.json")
    with open(json_path, "w") as f:
        json.dump(best_results, f, cls=NumpyEncoder, indent=2)

    return json_path


def extract_all(image_list, base_dir="input"):
    """Run OCR on all images, return list of generated JSON paths."""
    json_paths = []
    for img_name in image_list:
        print(f"  {img_name}")
        path = extract_text(img_name, base_dir)
        if path:
            json_paths.append(path)
    return json_paths


if __name__ == "__main__":
    input_dir = "input"
    valid_ext = ('.jpg', '.jpeg', '.png')
    if os.path.isdir(input_dir):
        images = sorted(f for f in os.listdir(input_dir) if f.lower().endswith(valid_ext))
    else:
        images = []
    if images:
        extract_all(images, base_dir=input_dir)
    else:
        print(f"No images found in {input_dir}/.")
