#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-21 08:23:30 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_base.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_base.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

wrap_function = None  # Will be populated later

def register_wrap_function(func):
    global wrap_function
    wrap_function = func

def get_wrapper(com_object, path=""):
    global wrap_function
    if wrap_function is None:
        from ._wrap import wrap
        wrap_function = wrap
    return wrap_function(com_object, path)

# EOF