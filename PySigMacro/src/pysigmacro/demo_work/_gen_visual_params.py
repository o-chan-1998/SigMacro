#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 19:14:29 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_visual_params.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_visual_params.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import pandas as pd

from ..data._create_padded_df import create_padded_df
from ..utils._calculate_nice_ticks import calculate_nice_ticks

# Parameters
# ------------------------------

def gen_visual_params(plot_type, n_cols=8, **kwargs):
    # Base Parameters
    out_dict = {}

    # Plot-type-specific Parameters
    PLOT_SPECIFIC_PARAMS = {
        "area": _gen_demo_visual_params_area(),
        "bar": _gen_demo_visual_params_bar(),
        "barh": _gen_demo_visual_params_barh(),
        "box": _gen_demo_visual_params_box(),
        "boxh": _gen_demo_visual_params_boxh(),
        "line": _gen_demo_visual_params_line(),
        "polar": _gen_demo_visual_params_polar(),
        "scatter": _gen_demo_visual_params_scatter(),
        "violin": _gen_demo_visual_params_violin(),
        # Special
        "filled_line": _gen_demo_visual_params_filled_line(),
        "contour": _gen_demo_visual_params_contour(),
        "conf_mat": _gen_demo_visual_params_conf_mat(),
    }[plot_type]
    out_dict.update(PLOT_SPECIFIC_PARAMS)

    # Passed parameters
    out_dict.update(kwargs)

    # Reformat
    try:
        xticks = dict(xticks=out_dict.pop("xticks"))
        yticks = dict(yticks=out_dict.pop("yticks"))
    except Exception as e:
        print(e)
        __import__("ipdb").set_trace()
    params_df = pd.DataFrame(pd.Series(out_dict)).reset_index()
    params_df.columns = ["visual parameter", "value"]

    # NaN padding
    params_df = create_padded_df(params_df, xticks, yticks)

    # Preserve additional columns for future expansion
    n_cols_preserve = n_cols - params_df.shape[1]
    for ii in range(n_cols_preserve):
        params_df[f"preserved {ii}"] = "NONE_STR"

    return params_df


def _gen_demo_visual_params_bar():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 45,
        "xmm": 40,
        "xscale": "category",
        "xmin": 0,
        "xmax": 11,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 16,
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_barh():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 21,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "category",
        "ymin": 0,
        "ymax": "NONE_STR",
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_area():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 21,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 1.65,
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_box():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 90,
        "xmm": 40,
        "xscale": "category",
        "xmin": 0,
        "xmax": "NONE_STR",
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 77,
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_boxh():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 88,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "category",
        "ymin": "NONE_STR",
        "ymax": "NONE_STR",
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_line():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 11,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 16,
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_polar():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 16,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40,
        "yscale": "linear",
        "ymin": "NONE_STR",
        "ymax": "NONE_STR",
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_scatter():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 21,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 21,
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_violin():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "category",
        "xmin": 0,
        "xmax": "NONE_STR",
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 21,
        "yticks": ["auto"],
    }


# Special
# ------------------------------


def _gen_demo_visual_params_conf_mat():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "category",
        "xmin": "NONE_STR",
        "xmax": "NONE_STR",
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "category",
        "ymin": ["NONE_STR"],
        "ymax": ["NONE_STR"],
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_contour():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "category",
        "xmin": "NONE_STR",
        "xmax": "NONE_STR",
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "category",
        "ymin": ["NONE_STR"],
        "ymax": ["NONE_STR"],
        "yticks": ["auto"],
    }


def _gen_demo_visual_params_filled_line():
    return {
        "xlabel": "X-Axis Label",
        "xrot": 0,
        "xmm": 40,
        "xscale": "linear",
        "xmin": 0,
        "xmax": 21,
        "xticks": ["auto"],
        "ylabel": "Y-Axis Label",
        "yrot": 0,
        "ymm": 40 * 0.7,
        "yscale": "linear",
        "ymin": 0,
        "ymax": 21,
        "yticks": ["auto"],
    }

# EOF