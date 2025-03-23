#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 10:40:04 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_crop_images.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_crop_images.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import cv2
import numpy as np
from PIL import Image
import glob

def find_content_area(image_path):
    # Check if it's a GIF file
    if image_path.lower().endswith('.gif'):
        try:
            # Use PIL for GIF files
            with Image.open(image_path) as img:
                # Convert to RGB for consistent processing
                img = img.convert('RGB')
                # Convert to numpy array for processing
                img_array = np.array(img)

                # Convert to grayscale
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)

                # Threshold the image to get all non-white areas
                _, thresh = cv2.threshold(gray, 254, 255, cv2.THRESH_BINARY_INV)

                # Find all non-zero points
                points = cv2.findNonZero(thresh)

                # Get the bounding rectangle
                if points is not None:
                    x, y, w, h = cv2.boundingRect(points)
                    return x, y, w, h
                else:
                    # If no points found, return the whole image
                    h, w = img_array.shape[:2]
                    return 0, 0, w, h
        except Exception as e:
            print(f"Error processing GIF with PIL: {e}")
            # Fall back to default dimensions
            return 0, 0, 100, 100

    # For non-GIF files, use OpenCV
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"Unable to read image file: {image_path}")

    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Threshold the image to get all non-white areas.
    # Set the threshold close to 255 to exclude only white or nearly white areas.
    _, thresh = cv2.threshold(gray, 254, 255, cv2.THRESH_BINARY_INV)

    # Find all non-zero points (colored or non-white pixels)
    points = cv2.findNonZero(thresh)

    # Get the bounding rectangle around the non-zero points
    if points is not None:
        x, y, w, h = cv2.boundingRect(points)
        return x, y, w, h
    else:
        # If no points found, return the whole image
        h, w = img.shape[:2]
        return 0, 0, w, h

def crop_image(input_path, margin=30, output_path=None):
    # Determine output path if not provided
    if output_path is None:
        basename = os.path.splitext(input_path)[0]
        ext = os.path.splitext(input_path)[1]
        output_path = f"{basename}_cropped{ext}"

    # Check if it's a GIF file
    if input_path.lower().endswith('.gif'):
        try:
            # Use PIL for GIF files
            with Image.open(input_path) as img:
                # Find content area
                x, y, w, h = find_content_area(input_path)

                # Calculate coordinates with margin
                x_start = max(x - margin, 0)
                y_start = max(y - margin, 0)
                x_end = min(x + w + margin, img.width)
                y_end = min(y + h + margin, img.height)

                # Crop the image
                cropped_img = img.crop((x_start, y_start, x_end, y_end))

                # Save the cropped image
                cropped_img.save(output_path)
                print(f"\n{input_path} was cropped and saved as {output_path}")
                return
        except Exception as e:
            print(f"Error processing GIF with PIL: {e}")
            # Continue with OpenCV method as fallback

    # For non-GIF files, use OpenCV
    img = cv2.imread(input_path)
    if img is None:
        raise FileNotFoundError(f"Unable to read image file: {input_path}")

    # Find the content area
    x, y, w, h = find_content_area(input_path)

    # Calculate the coordinates with margin, clamping to the image boundaries
    x_start = max(x - margin, 0)
    y_start = max(y - margin, 0)
    x_end = min(x + w + margin, img.shape[1])
    y_end = min(y + h + margin, img.shape[0])

    # Crop the image using the bounding rectangle with margin
    cropped_img = img[y_start:y_end, x_start:x_end]

    # Save the cropped image
    cv2.imwrite(output_path, cropped_img)
    print(f"\n{input_path} was cropped and saved as {output_path}")

def crop_tif(lpath_tif, margin=30):
    # This is a wrapper for backward compatibility
    # Read the image
    img = cv2.imread(lpath_tif)
    if img is None:
        raise FileNotFoundError(f"Unable to read image file: {lpath_tif}")

    # Find the content area
    x, y, w, h = find_content_area(lpath_tif)

    # Calculate the coordinates with margin, clamping to the image boundaries
    x_start = max(x - margin, 0)
    y_start = max(y - margin, 0)
    x_end = min(x + w + margin, img.shape[1])
    y_end = min(y + h + margin, img.shape[0])

    # Crop the image using the bounding rectangle with margin
    cropped_img = img[y_start:y_end, x_start:x_end]

    # Save the cropped image
    cv2.imwrite(lpath_tif, cropped_img)
    print(f"\n{lpath_tif} was cropped.")

if __name__ == "__main__":
    import argparse

    # Set up argument parser
    parser = argparse.ArgumentParser(description='Crop images to content area with margin')
    parser.add_argument('input_paths', nargs='+', help='Path(s) to input image(s), accepts wildcards')
    parser.add_argument(
        "--margin", type=int, default=30, help="Margin size around the content area."
    )
    parser.add_argument(
        "--output_dir", type=str, help="Directory to save cropped images (default: same as input)"
    )

    args = parser.parse_args()

    # Expand file paths (in case wildcards were used)
    input_files = []
    for path in args.input_paths:
        if os.path.isfile(path):
            input_files.append(path)
        else:
            # Use glob to expand wildcards
            expanded = glob.glob(path)
            input_files.extend(expanded)

    if not input_files:
        print("No input files found!")
    else:
        for input_path in input_files:
            try:
                # Determine output path
                if args.output_dir:
                    # Ensure output directory exists
                    os.makedirs(args.output_dir, exist_ok=True)

                    # Create output path in specified directory
                    filename = os.path.basename(input_path)
                    basename, ext = os.path.splitext(filename)
                    output_path = os.path.join(args.output_dir, f"{basename}_cropped{ext}")
                else:
                    # Default: save in same directory as input
                    basename, ext = os.path.splitext(input_path)
                    output_path = f"{basename}_cropped{ext}"

                # Crop the image
                crop_image(input_path, margin=args.margin, output_path=output_path)
            except Exception as e:
                print(f"Error processing {input_path}: {e}")

# EOF