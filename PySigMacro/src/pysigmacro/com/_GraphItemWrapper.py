#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 20:09:34 (ywatanabe)"
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

    def export_as_bmp(self, path=None, crop=False):
        """Export graph item as BMP, with optional cropping"""
        # Use self.path as default if path is not provided
        if path is None:
            path = os.path.splitext(self.path)[0] + ".bmp"
        # Export to BMP
        self._com_object.Export(path, "BMP")
        # Crop if requested
        if crop:
            from ..image import crop_images

            crop_images([path])
        return path

    def export_as_tif(self, path=None, crop=True, convert_from_bmp=True):
        """
        Export graph item as TIF, with optional cropping
        If convert_from_bmp is True, will first export as BMP then convert to TIFF
        (useful when direct TIFF export doesn't work properly)
        """
        # Use self.path as default if path is not provided
        if path is None:
            try:
                path = os.path.splitext(self.path)[0] + ".tif"
            except AttributeError:
                raise ValueError("No path provided and self.path is not set")

        if convert_from_bmp:
            # Create temporary BMP path
            bmp_path = os.path.splitext(path)[0] + ".bmp"
            # Export to BMP first
            self.export_as_bmp(bmp_path, crop=crop)
            # Convert to TIFF
            from ..image import convert_bmp_to_tiff

            tiff_path = convert_bmp_to_tiff(
                bmp_path
                if not crop
                else bmp_path.replace(".bmp", "_cropped.bmp")
            )
            # Delete temporary BMP if it was not cropped
            if not crop:
                try:
                    os.remove(bmp_path)
                except:
                    pass
            return tiff_path
        else:
            # Direct TIFF export
            try:
                self._com_object.Export(path, "TIF")
                # Crop if requested
                if crop:
                    from ..image import crop_images

                    crop_images([path])
                return path
            except Exception as e:
                print(f"Direct TIFF export failed: {e}")
                print("Attempting BMP conversion method...")
                return self.export_as_tif(
                    path, crop=crop, convert_from_bmp=True
                )


# class GraphItemWrapper(BaseCOMWrapper):
#     """Specialized wrapper for GraphItem object"""

#     __classname__ = "GraphItemWrapper"

#     def export_as_jpg(self, path, crop=False):
#         """Export graph item as JPG, with optional cropping"""
#         # Export to JPG
#         self._com_object.Export(path, "JPG")

#         # Crop if requested
#         if crop:
#             from ..image import crop_images
#             crop_images([path])

#         return path

#     def export_as_bmp(self, path, crop=False):
#         """Export graph item as BMP, with optional cropping"""
#         # Export to BMP
#         self._com_object.Export(path, "BMP")

#         # Crop if requested
#         if crop:
#             from ..image import crop_images
#             crop_images([path])

#         return path

#     def export_as_tif(self, path, crop=False, convert_from_bmp=False):
#         """
#         Export graph item as TIF, with optional cropping

#         If convert_from_bmp is True, will first export as BMP then convert to TIFF
#         (useful when direct TIFF export doesn't work properly)
#         """
#         if convert_from_bmp:
#             # Create temporary BMP path
#             import os
#             bmp_path = os.path.splitext(path)[0] + ".bmp"

#             # Export to BMP first
#             self.export_as_bmp(bmp_path, crop=crop)

#             # Convert to TIFF
#             from ..image import convert_bmp_to_tiff
#             tiff_path = convert_bmp_to_tiff(bmp_path if not crop else
#                                              bmp_path.replace(".bmp", "_cropped.bmp"))

#             # Delete temporary BMP if it was not cropped
#             if not crop:
#                 try:
#                     os.remove(bmp_path)
#                 except:
#                     pass

#             return tiff_path
#         else:
#             # Direct TIFF export
#             try:
#                 self._com_object.Export(path, "TIF")

#                 # Crop if requested
#                 if crop:
#                     from ..image import crop_images
#                     crop_images([path])

#                 return path
#             except Exception as e:
#                 print(f"Direct TIFF export failed: {e}")
#                 print("Attempting BMP conversion method...")
#                 return self.export_as_tif(path, crop=crop, convert_from_bmp=True)

# EOF