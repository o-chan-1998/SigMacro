#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 18:28:19 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_gen_graph_wizard_params.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_gen_graph_wizard_params.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np
import pandas as pd
from ._create_padded_df import create_padded_df
from ..const._PLOT_PARAMS_BASE import PLOT_PARAMS_BASE
from ..const import BGRA, COLORS


# Functions
# ------------------------------
def gen_graph_wizard_params_bar():
    return {
        "plot_type": "Vertical Bar Chart",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "None",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": "None",
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_barh():
    return {
        "plot_type": "Horizontal Bar Chart",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "YX Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "None",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_area():
    return {
        "plot_type": "Area Plot",
        "plot_style_type": "Simple Area",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_box():
    return {
        "plot_type": "Box Plot",
        "plot_style_type": "Vertical Box Plot",
        "plot_data_type": "X Many Y",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_boxh():
    return {
        "plot_type": "Box Plot",
        "plot_style_type": "Horizontal Box Plot",
        "plot_data_type": "Y Many X",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_line():
    return {
        "plot_type": "Line and Scatter Plot",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "None",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_filled_line():
    return {
        "plot_type": "Filled Line Plot",
        "plot_style_type": "Simple Error Bars",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "None",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_scatter():
    return {
        "plot_type": "Scatter Plot",
        "plot_style_type": "Simple Scatter",
        "plot_data_type": "XY Pair",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Standard Deviation",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "None",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_polar():
    return {
        "plot_type": "Polar Plot",
        "plot_style_type": "Lines",
        "plot_data_type": "Theta R",
        "plot_columns_per_plot": "ColumnsPerPlot",
        "plot_plot_columns_count_array": "PlotColumnCountArray",
        "plot_data_source": "Worksheet Columns",
        "plot_polarunits": "Radians",
        "plot_anguleunits": "Degrees",
        "plot_min_angle_row": 0,
        "plot_max_angle_row": 360,
        "plot_unknown1": np.nan,
        "plot_group_style": "Standard Deviation",
        "plot_use_automatic_legends": True,
    }

def gen_graph_wizard_params_violin():
    pass

def gen_graph_wizard_params(plot_type, **kwargs):
    # Plottype specific Parameters
    gen_graph_wizard_func = {
        "bar": gen_graph_wizard_params_bar,
        "barh": gen_graph_wizard_params_barh,
        "area": gen_graph_wizard_params_area,
        "box": gen_graph_wizard_params_box,
        "boxh": gen_graph_wizard_params_boxh,
        "line": gen_graph_wizard_params_line,
        "filled_line": gen_graph_wizard_params_filled_line,
        "polar": gen_graph_wizard_params_polar,
        "scatter": gen_graph_wizard_params_scatter,
        "violin": gen_graph_wizard_params_violin,
    }[plot_type]

    graph_wizard_params =  gen_graph_wizard_func()

    # Reformat
    params_df = pd.DataFrame(pd.Series(graph_wizard_params)).reset_index()
    params_df.columns = ["graph wizard parameter", "value"]

    return create_padded_df(params_df)

# EOF