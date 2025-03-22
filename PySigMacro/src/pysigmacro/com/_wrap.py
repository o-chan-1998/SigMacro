#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-21 09:44:57 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._COMWrapper import COMWrapper
from ._NotebooksWrapper import NotebooksWrapper
from ._NotebookItemsWrapper import NotebookItemsWrapper
from ._base import register_wrap_function

def wrap(com_object, path=""):
    """Factory function to create appropriate wrapper"""
    if path.endswith("Notebooks"):
        return NotebooksWrapper(com_object, path)
    if path.endswith("NotebookItems"):
        return NotebookItemsWrapper(com_object, path)
    else:
        return COMWrapper(com_object, path)

# Register the wrap function
register_wrap_function(wrap)

# EOF