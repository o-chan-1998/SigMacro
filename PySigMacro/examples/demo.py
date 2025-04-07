#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-06 01:49:54 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
if "SIGMACRO_JNB_PATH" not in os.environ:
    os.environ["SIGMACRO_JNB_PATH"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB"
if "SIGMACRO_TEMPLATES_DIR" not in os.environ:
    os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\templates"
if "SIGMAPLOT_BIN_PATH_WIN" not in os.environ:
    os.environ["SIGMAPLOT_BIN_PATH_WIN"] = rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

import pysigmacro as psm

# UNABAILABLE_PLOT_TYPES = ["violin", "contour", "conf_mat", "filled_line"]

# plot_types = ["scatter", "line"]
# plot_types = psm.const.PLOT_TYPES

for plot_type in psm.const.PLOT_TYPES:
    if plot_type != "scatter":
        continue
    plot_types = [plot_type for _ in range(13)]
    psm.demo.gen_csv(plot_types, save=True)
    psm.demo.gen_jnb(plot_types)

# try:
#     # CSV data
#     psm.demo.gen_csv(plot_types, save=True)
# except Exception as e:
#     print(f"Creating csv data for {plot_types} failed")
#     print(e)

# try:
#     # JNB and Figures
#     psm.demo.gen_jnb(plot_types)
# except Exception as e:
#     print(f"Creating JNB file for {plot_types} failed")
#     print(e)

# EOF