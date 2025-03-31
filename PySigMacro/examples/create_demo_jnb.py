#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-31 20:39:48 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_jnb.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_jnb.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables only set if not already defined
if "SIGMACRO_JNB_PATH" not in os.environ:
    os.environ["SIGMACRO_JNB_PATH"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro.JNB"
if "SIGMACRO_TEMPLATES_DIR" not in os.environ:
    # os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\SigMacro\SigMacro\Templates"
    os.environ["SIGMACRO_TEMPLATES_DIR"] = rf"C:\Users\{os.getlogin()}\Documents\Sigmacro\PySigMacro\src\pysigmacro\data\templates"
if "SIGMAPLOT_BIN_PATH_WIN" not in os.environ:
    os.environ["SIGMAPLOT_BIN_PATH_WIN"] = rf"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"


import pysigmacro as psm

for plot_type in psm.const.PLOT_TYPES:
    try:
        psm.data.create_templates(plot_type)
    except Exception as e:
        print(f"Creating template for {plot_type} failed")
        print(e)

# PLOT_TYPES_NOT_WORKING = [
#     "filled_line",
#     "polar",
#     "violin",
# ]

# Main
psm.data.create_templates(plot_type)

# EOF