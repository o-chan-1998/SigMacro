#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 11:14:31 (ywatanabe)"
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
from PIL import Image, ImageOps
import glob

def get_image_dpi(img):
    """Get the DPI of an image, defaulting to 96 if not available."""
    try:
        dpi_x, dpi_y = img.info.get('dpi', (96, 96))
        # Some images might return 0 for DPI
        return max(dpi_x, 1)  # Ensure DPI is at least 1
    except:
        return 96

def mm_to_pixels(mm, dpi):
    """Convert millimeters to pixels at the given DPI."""
    inches = mm / 25.4  # 1 inch = 25.4 mm
    return int(inches * dpi)

def _find_content_area(image_path):
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

def crop_image(input_path, margin=30):
    # Check if it's a GIF file
    if input_path.lower().endswith('.gif'):
        try:
            # Use PIL for GIF files
            with Image.open(input_path) as img:
                # Check if GIF has multiple frames
                is_animated = hasattr(img, 'n_frames') and img.n_frames > 1

                # Find content area
                x, y, w, h = _find_content_area(input_path)

                # Calculate coordinates with margin
                x_start = max(x - margin, 0)
                y_start = max(y - margin, 0)
                x_end = min(x + w + margin, img.width)
                y_end = min(y + h + margin, img.height)

                if is_animated:
                    # Process all frames for animated GIFs
                    frames = []
                    for frame_idx in range(img.n_frames):
                        img.seek(frame_idx)
                        frame_copy = img.copy()
                        cropped_frame = frame_copy.crop((x_start, y_start, x_end, y_end))
                        frames.append(cropped_frame)

                    # Return the first frame for now (we'll handle animation later)
                    return frames[0]
                else:
                    # Crop the image (single frame)
                    cropped_img = img.crop((x_start, y_start, x_end, y_end))
                    return cropped_img
        except Exception as e:
            print(f"Error processing GIF with PIL: {e}")
            # Continue with OpenCV method as fallback

    # For non-GIF files, use OpenCV
    img = cv2.imread(input_path)
    if img is None:
        raise FileNotFoundError(f"Unable to read image file: {input_path}")

    # Find the content area
    x, y, w, h = _find_content_area(input_path)

    # Calculate the coordinates with margin, clamping to the image boundaries
    x_start = max(x - margin, 0)
    y_start = max(y - margin, 0)
    x_end = min(x + w + margin, img.shape[1])
    y_end = min(y + h + margin, img.shape[0])

    # Crop the image using the bounding rectangle with margin
    cropped_img = img[y_start:y_end, x_start:x_end]

    # Convert OpenCV image to PIL Image
    pil_img = Image.fromarray(cv2.cvtColor(cropped_img, cv2.COLOR_BGR2RGB))
    return pil_img

def save_image(image, output_path):
    # Check if output path has extension
    base, ext = os.path.splitext(output_path)
    if not ext:
        # Default to PNG if no extension
        output_path = f"{base}.png"

    # Save the image
    image.save(output_path)
    print(f"\nImage saved as {output_path}")

def add_margins(image, width=None, height=None, width_mm=None, height_mm=None):
    # Get the DPI of the original image
    dpi = get_image_dpi(image)

    # Convert mm to pixels if provided
    if width_mm is not None:
        width = mm_to_pixels(width_mm, dpi)
    if height_mm is not None:
        height = mm_to_pixels(height_mm, dpi)

    def _add_margin_up_to_width(image, width=None):
        if width is not None:
            image_width = image.width
            if image_width < width:
                left_margin = (width - image_width) // 2
                right_margin = width - image_width - left_margin
                return ImageOps.expand(image, (left_margin, 0, right_margin, 0), fill=(255, 255, 255, 1))
        return image

    def _add_margin_up_to_height(image, height=None):
        if height is not None:
            image_height = image.height
            if image_height < height:
                top_margin = (height - image_height) // 2
                bottom_margin = height - image_height - top_margin
                return ImageOps.expand(image, (0, top_margin, 0, bottom_margin), fill=(255, 255, 255, 1))
        return image

    # Ensure image is in RGBA mode for transparent margins
    if image.mode != 'RGBA':
        image = image.convert('RGBA')

    image = _add_margin_up_to_width(image, width)
    image = _add_margin_up_to_height(image, height)

    return image

def main(args):
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
                    # Remove previous '_cropped' suffixes
                    basename = basename.replace("_cropped", "")
                    output_path = os.path.join(args.output_dir, f"{basename}_cropped{ext}")
                else:
                    # Default: save in same directory as input
                    basename, ext = os.path.splitext(input_path)
                    # Remove previous '_cropped' suffixes
                    basename = basename.replace("_cropped", "")
                    output_path = f"{basename}_cropped{ext}"

                if os.path.exists(output_path) and (not args.force):
                    print(f"Output path: {output_path} exists. Skipping...")
                    continue

                # Crop the image
                cropped_img = crop_image(input_path, margin=args.margin)

                # Add margins if dimensions specified
                if args.width or args.height or args.width_mm or args.height_mm:
                    cropped_img = add_margins(
                        cropped_img,
                        width=args.width,
                        height=args.height,
                        width_mm=args.width_mm,
                        height_mm=args.height_mm
                    )
                    # Get the DPI for reporting
                    dpi = get_image_dpi(cropped_img)
                    # Calculate actual dimensions for reporting
                    width_px = cropped_img.width
                    height_px = cropped_img.height
                    print(f"Adding margins to meet target dimensions: width={width_px}px, height={height_px}px (DPI: {dpi})")

                # Save the processed image
                save_image(cropped_img, output_path)

            except Exception as e:
                print(f"Error processing {input_path}: {e}")

if __name__ == "__main__":
    import argparse

    # Set up argument parser
    parser = argparse.ArgumentParser(description='Crop images to content area with margin')
    parser.add_argument('input_paths', nargs='+', help='Path(s) to input image(s), accepts wildcards')
    parser.add_argument(
        "--force", action="store_true", default=False, help="If true remove output path if exists."
    )
    parser.add_argument(
        "--margin", type=int, default=30, help="Margin pixel size around the content area."
    )
    parser.add_argument(
        "--width", type=int, help="Target width pixel size for the image (adds margins if needed)"
    )
    parser.add_argument(
        "--height", type=int, help="Target height pixel size for the image (adds margins if needed)"
    )
    parser.add_argument(
        "--width_mm", type=float, help="Target width in millimeters (adds margins if needed)"
    )
    parser.add_argument(
        "--height_mm", type=float, help="Target height in millimeters (adds margins if needed)"
    )
    parser.add_argument(
        "--dpi", type=int, default=300, help="DPI to use for mm to pixel conversion (default: 300)"
    )
    parser.add_argument(
        "--output_dir", type=str, help="Directory to save cropped images (default: same as input)"
    )
    args = parser.parse_args()
    main(args)

# EOF