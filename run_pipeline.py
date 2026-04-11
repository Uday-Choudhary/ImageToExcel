#!/usr/bin/env python3
"""
ImageToExcel CLI Pipeline

Entry point for processing images via Vision AI or Legacy OCR.
Uses the shared core modules and extractor strategies.

Usage:
    python run_pipeline.py                     # Vision mode (default)
    python run_pipeline.py --method ocr        # Legacy OCR mode
    python run_pipeline.py img1.jpg img2.png   # Specific images
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time

from dotenv import load_dotenv

from core.constants import INPUT_DIR, OUTPUT_DIR, VALID_IMAGE_EXTENSIONS
from core.excel_builder import build_excel_from_ocr, build_excel_from_vision

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# Load environment variables (override existing ones from shell to prevent bad caching)
load_dotenv(override=True)


def find_images() -> list[str]:
    """Find all valid image files in the input directory.

    Returns:
        Sorted list of image filenames.
    """
    if not os.path.isdir(INPUT_DIR):
        logger.warning("No '%s/' folder found. Create it and add your images.", INPUT_DIR)
        return []
    return sorted(
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(VALID_IMAGE_EXTENSIONS)
    )


def run_vision_pipeline(images: list[str]) -> None:
    """Execute the Vision AI extraction pipeline.

    Args:
        images: List of image filenames to process.
    """
    from extractors.vision_extractor import VisionExtractor

    extractor = VisionExtractor(save_json=True)
    logger.info("[1/2] Processing %d images with Vision AI...", len(images))

    results: list[tuple[str, dict]] = []
    for img_name in images:
        full_path = os.path.join(INPUT_DIR, img_name)
        data = extractor.extract_from_image(full_path, filename=img_name)
        if data:
            sheet_name = img_name.rsplit(".", 1)[0]
            results.append((sheet_name, data))

    if not results:
        logger.error("No data extracted from any image.")
        return

    logger.info("[2/2] Generating Excel workbook...")
    excel_bytes = build_excel_from_vision(results)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "Extracted_Data_Vision.xlsx")
    with open(output_path, "wb") as f:
        f.write(excel_bytes)
    logger.info("Saved %d sheet(s) → %s", len(results), output_path)


def run_ocr_pipeline(images: list[str]) -> None:
    """Execute the Legacy OCR extraction pipeline.

    Args:
        images: List of image filenames to process.
    """
    from extractors.ocr_extractor import OCRExtractor

    extractor = OCRExtractor()
    logger.info("[1/2] Processing %d images with Legacy OCR...", len(images))

    results: list[dict] = []
    for img_name in images:
        full_path = os.path.join(INPUT_DIR, img_name)
        data = extractor.extract_from_image(full_path, filename=img_name)
        if data:
            results.append(data)

    if not results:
        logger.error("No data extracted from any image.")
        return

    logger.info("[2/2] Generating Excel workbook...")
    excel_bytes = build_excel_from_ocr(results)

    if excel_bytes:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(OUTPUT_DIR, "Extracted_Data_OCR.xlsx")
        with open(output_path, "wb") as f:
            f.write(excel_bytes)
        logger.info("Saved → %s", output_path)
    else:
        logger.error("No structured data found to export.")


def main() -> None:
    """CLI entry point with argument parsing."""
    parser = argparse.ArgumentParser(
        description="ImageToExcel — Convert document images to structured Excel files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "images", nargs="*",
        help="Specific image files to process (default: all images in input/)",
    )
    parser.add_argument(
        "--method", choices=["vision", "ocr"], default="vision",
        help="Extraction method: 'vision' (recommended) or 'ocr' (offline legacy)",
    )

    args = parser.parse_args()
    start = time.time()

    images = args.images if args.images else find_images()
    if not images:
        logger.error("No images found in '%s/' or provided as arguments.", INPUT_DIR)
        sys.exit(1)

    logger.info("Starting pipeline (method=%s) — %d image(s)", args.method, len(images))

    if args.method == "vision":
        run_vision_pipeline(images)
    else:
        run_ocr_pipeline(images)

    elapsed = time.time() - start
    logger.info("Completed in %.1fs", elapsed)


if __name__ == "__main__":
    main()
