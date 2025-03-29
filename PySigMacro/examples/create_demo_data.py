#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-29 21:05:54 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/create_demo_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

# Environmental variables
os.environ["SIGMACRO_JNB_PATH"] = r"C:\Users\wyusu\Documents\SigMacro\SigMacro.JNB"
os.environ["SIGMACRO_TEMPLATES_DIR"] = r"C:\Users\wyusu\Documents\SigMacro\SigMacro\Templates"
os.environ["SIGMAPLOT_BIN_PATH_WIN"] = r"C:\Program Files (x86)\SigmaPlot\SPW16\Spw.exe"

import pysigmacro as psm
import pandas as pd
import numpy as np

# Demo data generation
# --------------------------------------------------

def create_demo_df(plot_type):
    # Parameters
    demo_params = psm.utils.gen_plot_params(plot_type)

    # Data
    demo_data = psm.utils.gen_plot_data(plot_type)

    # To df
    dict_all = {}
    dict_all.update(demo_params)
    dict_all.update(demo_data)

    return psm.data.create_padded_df(dict_all, None)

def main(plot_type):
    # Main
    df = create_demo_df(plot_type)

    # Saving
    sdir = os.getenv("SIGMACRO_TEMPLATES_DIR", __DIR__)
    spath = psm.path.to_win(os.path.join(sdir, f"{plot_type}.csv"))
    df.to_csv(spath, index=False)
    print(f"Saved to: {spath}")


if __name__ == '__main__':
    # List of all available plot types
    plot_types = [
        "bar",
        "barh",
        "area",
        "box",
        "boxh",
        "line",
        "filled_line",
        "polar",
        "scatter",
        "violin"
    ]

    for plot_type in plot_types:
        main(plot_type)

# EOF