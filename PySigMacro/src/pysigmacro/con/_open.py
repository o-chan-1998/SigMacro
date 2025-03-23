#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 12:20:42 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_open.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/con/_open.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import win32com.client
import subprocess
from ..path import to_win, to_wsl
from ._close_all import close_all

def open(lpath=None, close_others=False):
    if close_others:
        close_all()

    # SigmaPlot bin path
    sp_bin_wsl="/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe"
    sp_bin_win=r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

    if lpath:
        # JNB file path
        lpath = os.path.abspath(lpath)
        lpath_win = to_win(lpath)
        lpath_wsl = to_wsl(lpath)

        # Call SigmaPlot with the file as argument
        for sp_bin in [sp_bin_wsl, sp_bin_win]:
            for lpath in [lpath_win, lpath_wsl]:
                try:
                    if os.path.exists(lpath):
                        subprocess.Popen([sp_bin, lpath])
                    break
                except Exception as e:
                    pass

    return win32com.client.Dispatch("SigmaPlot.Application")

# EOF