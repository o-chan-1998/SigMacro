#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 18:35:19 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_csv.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_csv.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
if "SIGMACRO_JNB_PATH" not in os.environ:
    os.environ["SIGMACRO_JNB_PATH"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB"
if "SIGMACRO_TEMPLATES_DIR" not in os.environ:
    os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro\Templates"
if "SIGMAPLOT_BIN_PATH_WIN" not in os.environ:
    os.environ["SIGMAPLOT_BIN_PATH_WIN"] = rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

import pysigmacro as psm
for plot_type in psm.const.PLOT_TYPES:
    psm.data.create_demo_csv(plot_type, save=True)

# EOF