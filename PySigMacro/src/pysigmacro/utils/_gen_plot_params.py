#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-29 21:05:32 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_gen_plot_params.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_gen_plot_params.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const._PLOT_PARAMS_BASE import PLOT_PARAMS_BASE

# Parameters
# ------------------------------
PARAMS_FOR_BAR = {
"xscale": "category",
"xmax": None,
"xticks": None,
}

PARAMS_FOR_BARH = {
"yscale": "category",
"ymax": None,
"yticks": None,
}

PARAMS_FOR_AREA = {
"fill": True,
"fill_alpha": 0.5,
"xmin": 0,
"xmax": 10,
}

PARAMS_FOR_BOX = {
"xscale": "category",
"showmeans": True,
"notch": False,
}

PARAMS_FOR_BOXH = {
"yscale": "category",
"showmeans": True,
"notch": False,
}

PARAMS_FOR_LINE = {
"marker": "o",
"linestyle": "-",
"linewidth": 2,
"xmin": 0,
"xmax": 10,
}

PARAMS_FOR_FILLED_LINE = {
"fill": True,
"fill_alpha": 0.3,
"linestyle": "-",
"linewidth": 2,
"xmin": 0,
"xmax": 10,
}

PARAMS_FOR_POLAR = {
"projection": "polar",
"marker": "o",
"linestyle": "-",
}

PARAMS_FOR_SCATTER = {
"marker": "o",
"linestyle": "",
"alpha": 0.7,
}

PARAMS_FOR_VIOLIN = {
"showmeans": True,
"showextrema": True,
"xscale": "category",
}

def gen_plot_params(plot_type, **kwargs):
    # Base Parameters
    out_dict = {}
    out_dict.update(PLOT_PARAMS_BASE)

    # Plottype specific Parameters
    PLOT_SPECIFIC_PARAMS = {
        "bar": PARAMS_FOR_BAR,
        "barh": PARAMS_FOR_BARH,
        "area": PARAMS_FOR_AREA,
        "box": PARAMS_FOR_BOX,
        "boxh": PARAMS_FOR_BOXH,
        "line": PARAMS_FOR_LINE,
        "filled_line": PARAMS_FOR_FILLED_LINE,
        "polar": PARAMS_FOR_POLAR,
        "scatter": PARAMS_FOR_SCATTER,
        "violin": PARAMS_FOR_VIOLIN,
    }[plot_type]

    out_dict.update(PLOT_SPECIFIC_PARAMS)

    # Passed parameters
    out_dict.update(kwargs)

    return out_dict

# EOF