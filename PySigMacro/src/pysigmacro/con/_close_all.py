#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:46:02 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_close_all.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import subprocess
import time

def close_all():
    """
    Force close all instances of SigmaPlot application.

    This function uses Windows' taskkill command to forcefully terminate
    all running SigmaPlot processes. It waits for 2 seconds after sending
    the kill command to ensure processes have time to terminate.

    Returns:
        None

    Raises:
        Prints a warning message if an Exception occurs, but does not raise it.
    """
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