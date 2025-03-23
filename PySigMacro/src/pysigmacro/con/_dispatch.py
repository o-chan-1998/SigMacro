#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 12:20:41 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_dispatch.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_dispatch.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import win32com.client

def dispatch():
    return win32com.client.Dispatch("SigmaPlot.Application")

# EOF