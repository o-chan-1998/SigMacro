#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 12:20:40 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

__THIS_FILE__ = (
    "/home/ywatanabe/proj/SigMacro/PySigMacro/src/pysigmacro/utils/_close_all.py"
)
__THIS_DIR__ = os.path.dirname(__THIS_FILE__)

import subprocess
import time

def close_all():
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "spw.exe"],
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        time.sleep(2)
    except Exception as e:
        print(f"Warning when closing SigmaPlot: {e}")

# EOF