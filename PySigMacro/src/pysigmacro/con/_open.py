#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 19:30:11 (ywatanabe)"
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
from ..com._wrap import wrap

# def open(lpath=None, close_others=False):
#     try:
#         if close_others:
#             close_all()
#         # SigmaPlot bin path
#         sp_bin_wsl="/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe"
#         sp_bin_win=r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"
#         if lpath:
#             # JNB file path
#             lpath = os.path.abspath(lpath)
#             lpath_win = to_win(lpath)
#             lpath_wsl = to_wsl(lpath)
#             # Call SigmaPlot with the file as argument
#             for sp_bin in [sp_bin_wsl, sp_bin_win]:
#                 for lpath in [lpath_win, lpath_wsl]:
#                     try:
#                         # print(sp_bin, lpath)
#                         if os.path.exists(lpath):
#                             subprocess.Popen([sp_bin, lpath])
#                             break
#                     except Exception as e:
#                         pass
#         sp = win32com.client.Dispatch("SigmaPlot.Application")
#         # spw = wrap(sp)
#         spw = wrap(sp, "SigmaPlot")
#         spw.path = lpath
#         return spw
#     except Exception as e:
#         print(f"Error opening SigmaPlot: {str(e)}")
#         return None

def open(lpath=None, close_others=False):
    try:
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
                for path in [lpath_win, lpath_wsl]:
                    try:
                        # print(sp_bin, path)
                        if os.path.exists(path):
                            subprocess.Popen([sp_bin, path])
                            break
                    except Exception as e:
                        pass

        sp = win32com.client.Dispatch("SigmaPlot.Application")
        spw = wrap(sp, "SigmaPlot")
        # spw = wrap(sp, "SigmaPlot")

        # Set the path attribute on the wrapper object
        if lpath:
            spw._path = lpath

        return spw
    except Exception as e:
        print(f"Error opening SigmaPlot: {str(e)}")
        return None

# def open(lpath=None, close_others=False):

#     if close_others:
#         close_all()

#     # SigmaPlot bin path
#     sp_bin_wsl="/mnt/c/Program Files (x86)/SigmaPlot/SPW16/Spw.exe"
#     sp_bin_win=r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

#     if lpath:
#         # JNB file path
#         lpath = os.path.abspath(lpath)
#         lpath_win = to_win(lpath)
#         lpath_wsl = to_wsl(lpath)

#         # Call SigmaPlot with the file as argument
#         for sp_bin in [sp_bin_wsl, sp_bin_win]:
#             for lpath in [lpath_win, lpath_wsl]:
#                 try:
#                     # print(sp_bin, lpath)
#                     if os.path.exists(lpath):
#                         subprocess.Popen([sp_bin, lpath])
#                     break
#                 except Exception as e:
#                     pass

#     sp = win32com.client.Dispatch("SigmaPlot.Application")
#     spw = wrap(sp)
#     spw.path = lpath
#     return spw

# EOF