#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-22 17:08:48 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookItemsWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebookItemsWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
import re
from ._COMWrapper import COMWrapper
from ._base import get_wrapper

class NotebookItemsWrapper(COMWrapper):
    """Specialized wrapper for NotebookItems collection"""

    def _list(self):
        """List all notebook items."""
        print("NotebookItems")
        for ii in range(self._com_object.Count):
            print(ii, self._com_object[ii].Name)

    def clear(self):
        """Clear notebook items with default naming pattern (e.g., Section 1, Section 2, ...)"""
        pattern = re.compile(r"^Section\s*\d+$")
        # Need to iterate in reverse because closing affects the collection indices
        for ii in range(self._com_object.Count - 1, -1, -1):
        # for ii in range(self._com_object.Count):
            print(ii, self._com_object[ii].Name)
            if pattern.match(self._com_object[ii].Name):
                try:
                    self._com_object[ii].Close(False)
                except IndexError as e:
                    print(f"Could not close notebook item {self._com_object[ii].Name}: {e}")

    def add_worksheet(self, name=None):
        """Add a new worksheet to the notebook"""
        worksheet_item = get_wrapper(self._com_object.Add(CT_WORKSHEET), self.access_path)
        if name:
            worksheet_item.Name = name
        return worksheet_item

    def add_graph(self, name=None):
        """Add a new graph to the notebook"""
        graph_item = get_wrapper(self._com_object.Add(CT_GRAPHICPAGE), self.access_path)
        if name:
            graph_item.Name = name
        return graph_item

    def add_report(self, name=None):
        """Add a new report to the notebook"""
        report_item = get_wrapper(self._com_object.Add(CT_REPORT), self.access_path)
        if name:
            report_item.Name = name
        return report_item

    def add_section(self, name=None):
        """Add a new section to the notebook"""
        section_item = get_wrapper(self._com_object.Add(CT_FOLDER), self.access_path)
        if name:
            section_item.Name = name
        return section_item

    def Item(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}({key})" if self._access_path else f"({key})"
        return get_wrapper(result, access_path)

    def __call__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}({key})" if self._access_path else f"({key})"
        return get_wrapper(result, access_path)

    def __getitem__(self, key):
        if key == -1 and hasattr(self._com_object, "Count"):
            key = self._com_object.Count - 1
        result = self._com_object(key)
        access_path = f"{self._access_path}[{key}]" if self._access_path else f"[{key}]"
        return get_wrapper(result, access_path)

    @property
    def list(self):
        return self._list()

# EOF