#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 19:40:43 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_WorksheetItemWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_WorksheetItemWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
from ._BaseCOMWrapper import BaseCOMWrapper

class WorksheetItemWrapper(BaseCOMWrapper):
    """Specialized wrapper for WorksheetItem object"""
    __classname__ = "WorksheetItemWrapper"

# EOF