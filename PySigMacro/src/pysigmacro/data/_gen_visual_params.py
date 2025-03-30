#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 18:39:40 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_gen_visual_params.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_gen_visual_params.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..const._PLOT_PARAMS_BASE import PLOT_PARAMS_BASE
from ._create_padded_df import create_padded_df
import numpy as np
import pandas as pd

# Parameters
# ------------------------------
PARAMS_FOR_BAR = {
    "xscale": "category",
    "xmax": np.nan,
    "xticks": [np.nan],
}

PARAMS_FOR_BARH = {
    "yscale": "category",
    "ymax": np.nan,
    "yticks": [np.nan],
}

PARAMS_FOR_AREA = {
}

PARAMS_FOR_BOX = {
    "xscale": "category",
    "xmax": np.nan,
    "xticks": [np.nan],
}

PARAMS_FOR_BOXH = {
    "yscale": "category",
    "ymax": np.nan,
    "yticks": [np.nan],
}

PARAMS_FOR_LINE = {
}

PARAMS_FOR_FILLED_LINE = {
}

PARAMS_FOR_POLAR = {
}

PARAMS_FOR_SCATTER = {
}

PARAMS_FOR_VIOLIN = {
    "xscale": "category",
    "xmax": np.nan,
    "xticks": [np.nan],
}

def gen_visual_params(plot_type, n_cols=8, **kwargs):
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

    # Reformat
    xticks = dict(xticks = out_dict.pop("xticks"))
    yticks = dict(yticks = out_dict.pop("yticks"))
    params_df = pd.DataFrame(pd.Series(out_dict)).reset_index()
    params_df.columns = ["visual parameter", "value"]

    # NaN padding
    params_df = create_padded_df(params_df, xticks, yticks)

    # Preserve additional columns for future expansion
    n_cols_preserve = n_cols - params_df.shape[1]
    for ii in range(n_cols_preserve):
        params_df[f"preserved {ii}"] = np.nan

    return params_df

# EOF