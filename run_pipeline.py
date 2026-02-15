#!/usr/bin/env python3
"""
Image to Excel Pipeline
Entry point for processing images via OCR or Vision models.
"""

import os
import sys
import time
import argparse
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# ─── Folder Configuration ────────────────────────────────────────────
INPUT_DIR = "input"


def find_images():
    """Find all source images in the input/ folder."""
    valid_ext = ('.jpg', '.jpeg', '.png')
    if not os.path.isdir(INPUT_DIR):
        print(f"No {INPUT_DIR}/ folder found. Create it and add your images.")
        return []
    return sorted(
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(valid_ext)
    )


def main():
    parser = argparse.ArgumentParser(description="Image to Excel Pipeline")
    parser.add_argument("images", nargs="*", help="Specific images to process")
    parser.add_argument("--method", choices=["ocr", "vision"], default="vision", help="Extraction method: 'vision' (Best) or 'ocr' (Legacy)")
    
    args = parser.parse_args()
    start = time.time()

    # Determine images to process
    images = args.images if args.images else find_images()

    if not images:
        print(f"No images found in {INPUT_DIR}/ or provided as arguments.")
        return

    print(f"Starting pipeline (Method: {args.method}) - Processing {len(images)} images")

    if args.method == "ocr":
        print("\n[1/3] Preprocessing...")
        from process_images import preprocess_all
        preprocess_all(images, base_dir=INPUT_DIR)

        print("\n[2/3] Running OCR...")
        from ocr_extraction import extract_all
        extract_all(images, base_dir=INPUT_DIR)

        print("\n[3/3] Generating Excel...")
        from convert_to_excel import save_to_excel
        save_to_excel(output_file="Extracted_Data_OCR.xlsx")

    elif args.method == "vision":
        print("\n[1/2] Processing with Vision Model...")
        from vision_processor import process_all_images
        process_all_images(images, base_dir=INPUT_DIR)

        print("\n[2/2] Generating Excel...")
        from json_to_excel import save_to_excel
        save_to_excel(output_file="Extracted_Data_Vision.xlsx")

    elapsed = time.time() - start
    print(f"\nCompleted in {elapsed:.1f}s")


if __name__ == "__main__":
    main()
