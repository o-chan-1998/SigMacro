#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-25 17:38:58 (ywatanabe)"
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

    def import_df(self, df):
        from ..data._import_data import import_data as ps_data_import_data
        self.datatable = ps_data_import_data(self, df)

# EOF