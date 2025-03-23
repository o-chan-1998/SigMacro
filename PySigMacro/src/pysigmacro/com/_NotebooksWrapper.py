#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-22 17:08:15 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebooksWrapper.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_NotebooksWrapper.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const import *
import re
from ._COMWrapper import COMWrapper
from ._base import get_wrapper

class NotebooksWrapper(COMWrapper):
    """Specialized wrapper for Notebooks collection"""
    def _list(self):
        """List all notebooks."""
        print("Notebooks")
        for ii in range(self._com_object.Count):
            print(ii, self._com_object[ii].Name)

    def clear(self):
        """Clear notebooks with default naming pattern (e.g., Notebook1, Notebook2, ...)"""
        pattern = re.compile(r"^Notebook\d+$")

        # Need to iterate in reverse because closing affects the collection indices
        for ii in range(self._com_object.Count - 1, -1, -1):
        # for ii in range(self._com_object.Count):
            if pattern.match(self._com_object[ii].Name):
                try:
                    self._com_object[ii].Close(False)
                except IndexError as e:
                    print(f"Could not close notebook {self._com_object[ii].Name}: {e}")

    def Add(self):
        """Add a new notebook and return the wrapped notebook object"""
        # Use the COM object's internal method to add a notebook
        try:
            # Different approach to access the Add method
            new_notebook = self._com_object.Add
            # access_path = f"{self._access_path}.Add()" if self._access_path else "Add()"
            access_path = self._access_path if self._access_path else ""
            return get_wrapper(new_notebook, access_path)
        except Exception as e:
            print(f"Error creating notebook: {e}")
            return None

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