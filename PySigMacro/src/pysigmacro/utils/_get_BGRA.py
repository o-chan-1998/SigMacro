#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-29 19:26:16 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_color.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_get_color.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const._COLORS import BGRA

def get_BGRA(color_str, alpha=1.0):
    rgba = RGBA[color_str]
    bgra = [rgba[2], rgba[1], rgba[0], rgba[3]]
    bgra[3] = alpha
    return bgra

# EOF