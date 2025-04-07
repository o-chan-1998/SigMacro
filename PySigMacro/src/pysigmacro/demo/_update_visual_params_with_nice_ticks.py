#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-05 05:54:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_update_visual_params_with_nice_ticks.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_update_visual_params_with_nice_ticks.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import pandas as pd
import numpy as np
from ..utils._calculate_nice_ticks import calculate_nice_ticks


def prefer_int(float_or_int_value):
    if float(float_or_int_value) == int(float_or_int_value):
        return int(float_or_int_value)
    elif abs(float(float_or_int_value) - round(float_or_int_value)) < 1e-10:
        # Handle floating point precision issues
        return int(round(float_or_int_value))
    else:
        return float(float_or_int_value)


def _extract_numeric_values(df, possible_numeric_columns):
    """Extract numeric values from dataframe columns"""
    numeric_data = []
    for i_col, col in enumerate(df.columns):
        try:
            if col in possible_numeric_columns:
                numeric_data.append(pd.to_numeric(df.iloc[:, i_col], errors="coerce"))
        except:
            pass

    if len(numeric_data) > 0:
        return pd.concat(numeric_data, axis=1)
    else:
        raise ValueError("Numeric data not found")


def _calculate_x_nice_ticks(
    df_visual_params, df_data, n_elems_in_chunk, pad_perc
):
    """Update x-axis ticks in visual parameters"""
    try:
        # Get numeric data
        possible_numeric_columns = ["x", "xerr", "x_lower", "x_upper", "theta"]
        x_values = _extract_numeric_values(df_data, possible_numeric_columns)

        if np.isnan(x_values).all().all():
            return ["NONE_STR"], "NONE_STR", "NONE_STR"

        # Calculate min/max and padded ranges
        x_min = np.nanmin(x_values.values)
        x_max = np.nanmax(x_values.values)

        x_length = x_max - x_min
        x_pad = x_length * pad_perc / 100.0
        x_padded_min = x_min - x_pad
        x_padded_max = x_max + x_pad

        # 0 is not padded
        if x_min == 0:
            x_padded_min = 0

        # Calculate nice ticks
        x_nice_ticks = calculate_nice_ticks(x_padded_min, x_padded_max)
        return x_nice_ticks, x_padded_min, x_padded_max

    except Exception as e:
        return ["NONE_STR"], "NONE_STR", "NONE_STR"


def _calculate_y_nice_ticks(
    df_visual_params, df_data, n_elems_in_chunk, pad_perc
):
    """Update y-axis ticks in visual parameters"""
    try:
        possible_numeric_columns = ["y", "yerr", "y_lower", "y_upper", "r"]

        # Get numeric data
        y_values = _extract_numeric_values(df_data, possible_numeric_columns)

        if np.isnan(y_values).all().all():
            return ["NONE_STR"], "NONE_STR", "NONE_STR"

        # Calculate min/max and padded ranges
        y_min = np.nanmin(y_values.values)
        y_max = np.nanmax(y_values.values)

        y_length = y_max - y_min
        y_pad = y_length * pad_perc / 100.0
        y_padded_min = y_min - y_pad
        y_padded_max = y_max + y_pad

        # 0 is not padded
        if y_min == 0:
            y_padded_min = 0

        # Calculate nice ticks
        y_nice_ticks = calculate_nice_ticks(y_padded_min, y_padded_max)

        return y_nice_ticks, y_padded_min, y_padded_max

    except Exception as e:
        print(f"Warning: Error calculating y-axis nice ticks: {e}")
        return ["NONE_STR"], "NONE_STR", "NONE_STR"


def _update_xticks(df_visual_params, x_nice_ticks, x_padded_min, x_padded_max):
    try:

        # Get the index of the value column
        param_values_col = (
            df_visual_params.columns.get_loc("visual parameter") + 1
        )
        xmin_row = df_visual_params.index[
            df_visual_params["visual parameter"] == "xmin"
        ].tolist()[0]
        xmax_row = df_visual_params.index[
            df_visual_params["visual parameter"] == "xmax"
        ].tolist()[0]

        df_visual_params.iloc[xmin_row, param_values_col] = prefer_int(
            x_padded_min
        )
        df_visual_params.iloc[xmax_row, param_values_col] = prefer_int(
            x_padded_max
        )

        # Update xticks
        for i, tick in enumerate(x_nice_ticks):
            if i < len(df_visual_params):
                df_visual_params.loc[i, "xticks"] = tick

    except Exception as e:
        print(f"Warning: Error calculating y-axis nice ticks: {e}")

    return df_visual_params


def _update_yticks(df_visual_params, y_nice_ticks, y_padded_min, y_padded_max):
    try:
        # Get the index of the value column
        param_values_col = (
            df_visual_params.columns.get_loc("visual parameter") + 1
        )
        ymin_row = df_visual_params.index[
            df_visual_params["visual parameter"] == "ymin"
        ].tolist()[0]
        ymax_row = df_visual_params.index[
            df_visual_params["visual parameter"] == "ymax"
        ].tolist()[0]

        # Update y-axis parameters
        df_visual_params.iloc[ymin_row, param_values_col] = prefer_int(
            y_padded_min
        )
        df_visual_params.iloc[ymax_row, param_values_col] = prefer_int(
            y_padded_max
        )

        # Update yticks
        for i, tick in enumerate(y_nice_ticks):
            if i < len(df_visual_params):
                df_visual_params.loc[i, "yticks"] = tick

    except Exception as e:
        print(f"Warning: Error calculating y-axis nice ticks: {e}")

    return df_visual_params


def update_visual_params_with_nice_ticks(
    plot_type, df_visual_params, df_data, n_elems_in_chunk=9, pad_perc=5
):
    # Nice Ticks when "auto" specified
    is_xticks_auto = df_visual_params["xticks"].iloc[0] == "auto"
    is_yticks_auto = df_visual_params["yticks"].iloc[0] == "auto"

    # Check if plot type is polar
    is_polar = plot_type == "polar"

    # Parameters for nice ticks calculation

    if not is_polar:
        # Update x-axis ticks if auto
        if is_xticks_auto:
            x_nice_ticks, x_padded_min, x_padded_max = _calculate_x_nice_ticks(
                df_visual_params, df_data, n_elems_in_chunk, pad_perc
            )
            df_visual_params = _update_xticks(
                df_visual_params, x_nice_ticks, x_padded_min, x_padded_max
            )

        # Update y-axis ticks if auto
        if is_yticks_auto:
            y_nice_ticks, y_padded_min, y_padded_max = _calculate_y_nice_ticks(
                df_visual_params, df_data, n_elems_in_chunk, pad_perc
            )
            df_visual_params = _update_yticks(
                df_visual_params, y_nice_ticks, y_padded_min, y_padded_max
            )

        return df_visual_params

    # For polar, x and y is inverse.
    # xticks: xaxis ticks | xvalues: degrees | yticks: yaxis ticks | yvalues: radius
    elif is_polar:
        if is_yticks_auto:
            y_padded_min = 0
            y_padded_max = 360
            y_nice_ticks = [0, 90, 180, 270]
            df_visual_params = _update_yticks(
                df_visual_params, y_nice_ticks, y_padded_min, y_padded_max
            )

        if is_xticks_auto:
            y_nice_ticks, y_padded_min, y_padded_max = _calculate_y_nice_ticks(
                df_visual_params, df_data, n_elems_in_chunk, pad_perc
            )
            x_nice_ticks = y_nice_ticks
            x_padded_min = y_padded_min
            x_padded_max = y_padded_max
            df_visual_params = _update_xticks(
                df_visual_params, x_nice_ticks, x_padded_min, x_padded_max
            )
        return df_visual_params

# EOF