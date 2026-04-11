"""
Image preprocessing pipeline for the OCR extraction path.

Applies deskewing, denoising, CLAHE contrast enhancement, and
sharpening to improve OCR accuracy on document images.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

import cv2
import numpy as np

from core.constants import DESKEW_MAX_ANGLE, DESKEW_MIN_ANGLE, PREPROCESSED_DIR

logger = logging.getLogger(__name__)


def deskew_image(image: np.ndarray) -> np.ndarray:
    """Detect and correct image skew using minAreaRect on text contours.

    Only corrects angles between ±0.5° and ±15° to avoid
    over-rotating non-skewed or severely rotated images.

    Args:
        image: Input image as a numpy array (grayscale or BGR).

    Returns:
        The deskewed image, or original if skew is negligible/extreme.
    """
    gray = image if len(image.shape) == 2 else cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) < 50:
        return image

    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    elif angle > 45:
        angle = -(angle - 90)
    else:
        angle = -angle

    if abs(angle) < DESKEW_MIN_ANGLE or abs(angle) > DESKEW_MAX_ANGLE:
        return image

    h, w = image.shape[:2]
    center = (w // 2, h // 2)
    rotation_matrix = cv2.getRotationMatrix2D(center, angle, 1.0)

    logger.debug("Deskewing by %.2f degrees", angle)
    return cv2.warpAffine(
        image, rotation_matrix, (w, h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )


def preprocess_image(image_path: str, output_path: str) -> Optional[np.ndarray]:
    """Apply full preprocessing pipeline to a single image.

    Pipeline: grayscale → deskew → denoise → CLAHE → sharpen.

    Args:
        image_path: Path to the input image.
        output_path: Path to write the preprocessed image.

    Returns:
        The preprocessed image array, or None if the image could not be read.
    """
    img = cv2.imread(image_path)
    if img is None:
        logger.error("Could not read image: %s", image_path)
        return None

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = deskew_image(gray)

    denoised = cv2.fastNlMeansDenoising(gray, h=10, templateWindowSize=7, searchWindowSize=21)

    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(denoised)

    sharpen_kernel = np.array([
        [0, -0.5, 0],
        [-0.5, 3, -0.5],
        [0, -0.5, 0],
    ])
    sharpened = cv2.filter2D(enhanced, -1, sharpen_kernel)
    result = np.clip(sharpened, 0, 255).astype(np.uint8)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    cv2.imwrite(output_path, result)
    logger.info("Preprocessed: %s → %s", os.path.basename(image_path), output_path)

    return result


def preprocess_all(image_list: list[str], base_dir: str = "input") -> list[str]:
    """Preprocess a batch of images, saving to the preprocessed directory.

    Args:
        image_list: List of image filenames.
        base_dir: Directory containing the source images.

    Returns:
        List of paths to successfully preprocessed images.
    """
    os.makedirs(PREPROCESSED_DIR, exist_ok=True)
    output_paths: list[str] = []

    for img_name in image_list:
        input_path = os.path.join(base_dir, img_name)
        output_path = os.path.join(PREPROCESSED_DIR, img_name)
        result = preprocess_image(input_path, output_path)
        if result is not None:
            output_paths.append(output_path)

    logger.info("Preprocessed %d / %d images", len(output_paths), len(image_list))
    return output_paths
