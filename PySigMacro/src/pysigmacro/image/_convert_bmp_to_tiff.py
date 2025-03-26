#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:48:05 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_convert_bmp_to_tiff.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/image/_convert_bmp_to_tiff.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from PIL import Image

def convert_bmp_to_tiff(bmp_path):
    """
    Convert BMP to TIFF with lossless compression

    This function takes a BMP file path, opens the image using Pillow,
    and saves it as a TIFF file with no compression and maximum quality.

    Args:
        bmp_path (str): Path to the input BMP file

    Returns:
        str: Path to the converted TIFF file
    """
    tiff_path = bmp_path.replace(".bmp", ".tiff")
    with Image.open(bmp_path) as img:
        img.save(tiff_path, format="TIFF", compression=None, quality=100)
    return tiff_path

if __name__ == '__main__':
    # Convert the BMP we just exported
    bmp_path = PATH.replace(".JNB", ".bmp")
    tiff_path = convert_bmp_to_tiff(bmp_path)
    print(f"Converted to TIFF: {tiff_path}")

# EOF