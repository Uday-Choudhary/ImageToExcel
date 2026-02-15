"""
Image Preprocessing
Enhances images for OCR using deskewing and denoising.
"""

import cv2
import numpy as np
import os

PREPROCESSED_DIR = "preprocessed"


def deskew_image(image):
    """Detect and correct image skew using minAreaRect on text contours."""
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

    if abs(angle) < 0.5 or abs(angle) > 15:
        return image

    h, w = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    return cv2.warpAffine(image, M, (w, h),
                          flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_REPLICATE)


def clean_image(image_path, output_path):
    """
    Preprocess image for better OCR accuracy:
    grayscale → deskew → denoise → CLAHE → sharpen.
    """
    img = cv2.imread(image_path)
    if img is None:
        print(f"  ERROR: Could not read {image_path}")
        return None

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = deskew_image(gray)

    denoised = cv2.fastNlMeansDenoising(gray, h=10,
                                         templateWindowSize=7,
                                         searchWindowSize=21)

    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(denoised)

    sharpen_kernel = np.array([[0, -0.5, 0],
                               [-0.5, 3, -0.5],
                               [0, -0.5, 0]])
    sharpened = cv2.filter2D(enhanced, -1, sharpen_kernel)
    result = np.clip(sharpened, 0, 255).astype(np.uint8)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    cv2.imwrite(output_path, result)
    print(f"  {os.path.basename(image_path)} → {output_path}")
    return result


def preprocess_all(image_list, base_dir="input"):
    """Preprocess a list of images, saving to preprocessed/ folder."""
    out_dir = os.path.join(PREPROCESSED_DIR)
    os.makedirs(out_dir, exist_ok=True)

    for img_name in image_list:
        input_path = os.path.join(base_dir, img_name)
        output_path = os.path.join(out_dir, img_name)
        clean_image(input_path, output_path)


if __name__ == "__main__":
    input_dir = "input"
    valid_ext = ('.jpg', '.jpeg', '.png')
    if os.path.isdir(input_dir):
        images = sorted(f for f in os.listdir(input_dir) if f.lower().endswith(valid_ext))
    else:
        images = []
    if images:
        print(f"Preprocessing {len(images)} images from {input_dir}/...")
        preprocess_all(images, base_dir=input_dir)
    else:
        print(f"No images found in {input_dir}/.")
