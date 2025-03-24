#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 17:18:04 (ywatanabe)"
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
    """Convert BMP to TIFF with lossless compression"""
    tiff_path = bmp_path.replace(".bmp", ".tiff")
    with Image.open(bmp_path) as img:
        # Save as TIFF with LZW compression (lossless)
        # img.save(tiff_path, format="TIFF", compression="tiff_lzw")
        img.save(tiff_path, format="TIFF", compression=None, quality=100)
    return tiff_path

if __name__ == '__main__':
    # Convert the BMP we just exported
    bmp_path = PATH.replace(".JNB", ".bmp")
    tiff_path = convert_bmp_to_tiff(bmp_path)
    print(f"Converted to TIFF: {tiff_path}")

# EOF