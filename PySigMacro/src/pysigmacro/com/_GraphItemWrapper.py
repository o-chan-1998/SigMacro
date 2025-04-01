#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 17:39:59 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphItemWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_GraphItemWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
from ._BaseCOMWrapper import BaseCOMWrapper
from ..utils._remove import remove


class GraphItemWrapper(BaseCOMWrapper):
    """Specialized wrapper for GraphItem object"""

    __classname__ = "GraphItemWrapper"

    def export_as_jpg(self, path=None, crop=False):
        """Export graph item as JPG, with optional cropping"""
        # Use self.path as default if path is not provided
        if path is None:
            path = os.path.splitext(self.path)[0] + ".jpg"
        # Export to JPG
        self._com_object.Export(path, "JPG")
        # Crop if requested
        if crop:
            from ..image import crop_images

            crop_images([path])
        return path

    def export_as_bmp(self, path=None, crop=False, keep_orig=True):
        """Export graph item as BMP, with optional cropping"""
        # Use self.path as default if path is not provided
        if path is None:
            path = os.path.splitext(self.path)[0] + ".bmp"
        # Export to BMP
        self._com_object.Export(path, "BMP")

        # Crop if requested
        if crop:
            from ..image import crop_images
            crop_images([path], keep_orig=keep_orig)

        return path

    def export_as_tif(self, path=None, crop=True, keep_orig=True):
        """
        Export graph item as TIF, with optional cropping
        If convert_from_bmp is True, will first export as BMP then convert to TIFF
        (useful when direct TIFF export doesn't work properly)
        """
        from ..image import convert_bmp_to_tiff

        # JNB to BMP
        if path is None:
            try:
                path = os.path.splitext(self.path)[0] + ".tif"
            except AttributeError:
                raise ValueError("No path provided and self.path is not set")
        bmp_path = os.path.splitext(path)[0] + ".bmp"
        self.export_as_bmp(bmp_path, crop=crop, keep_orig=keep_orig)

        # BMP to TIFF
        actual_bmp_path = bmp_path if not crop else bmp_path.replace(".bmp", "_cropped.bmp")
        tiff_path = convert_bmp_to_tiff(actual_bmp_path, keep_orig=keep_orig)

        return tiff_path

    def rename_xy_labels(self, xlabel, ylabel):
        from ..utils._run_macro import run_macro
        run_macro(
            self, "RenameXYLabels_macro", xlabel=xlabel, ylabel=ylabel
        )

# EOF