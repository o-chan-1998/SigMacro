#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 06:18:18 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np

from ..const import BGRA, BGRA_FAKE, COLORS
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params
from scipy import stats
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import to_rgba
from ._gen_data_heatmap import _gen_data_heatmap
from ._gen_single_data_violin import _gen_single_data_violin, _gen_single_data_violinh
from ._gen_single_data_filled_line import _gen_single_data_filled_line

# Main
# ------------------------------

def gen_data(plot_types):

    chunks_df = pd.DataFrame()
    for i_plot, plot_type in enumerate(plot_types):
        try:
            gen_single_data_func = {
                "scatter": _gen_single_data_scatter,
                "line": _gen_single_data_line,
                "line_yerr": _gen_single_data_line_yerr,
                "lines_y_many_x": _gen_single_data_lines_y_many_x,
                "lines_x_many_y": _gen_single_data_lines_x_many_y,
                "bar": _gen_single_data_bar,
                "barh": _gen_single_data_barh,
                "area": _gen_single_data_area,
                "box": _gen_single_data_box,
                "boxh": _gen_single_data_boxh,
                "polar": _gen_single_data_polar,
                "violin": _gen_single_data_violin,
                "violinh": _gen_single_data_violinh,
                "heatmap": _gen_data_heatmap,
                "filled_line": _gen_single_data_filled_line,
                "contour": _gen_single_data_contour,
            }[plot_type]

            if plot_type == "heatmap":
                chunks_df = gen_single_data_func(i_plot)

            if plot_type in ["violin", "violinh"]:
                single_chunk_df = gen_single_data_func(i_plot)
                chunks_df = create_padded_df(chunks_df, single_chunk_df)

            if plot_type == "filled_line":
                single_chunk_df = gen_single_data_func(i_plot)
                chunks_df = create_padded_df(chunks_df, single_chunk_df)

            else:
                single_plot_dict = gen_single_data_func(i_plot)
                single_plot_df = create_padded_df(single_plot_dict)
                gw_df = create_graph_wizard_params(plot_type, i_plot)
                single_chunk = create_padded_df(gw_df, single_plot_df)
                chunks_df = create_padded_df(chunks_df, single_chunk)
        except Exception as e:
            print(plot_type)
            print(e)

    return chunks_df


# Special
# ------------------------------
# def _gen_single_data_filled_line(ii, alpha=0.5):
#     # Random Seed
#     np.random.seed(ii * 555)
#     # X
#     x = np.linspace(0, 10, 20) + ii
#     # Y
#     y = np.exp(-((x - 5 * (ii % 3)) ** 2) / 10)
#     y_lower = y - np.random.normal(0, 0.05 * (ii + 1), size=len(x))
#     y_upper = y + np.random.normal(0, 0.05 * (ii + 1), size=len(x))

#     # Color
#     bgra = BGRA[COLORS[ii % len(COLORS)]]
#     bgra[-1] = alpha
#     return dict(
#         x=x,
#         y_lower=y_lower,
#         y=y,
#         y_upper=y_upper,
#         bgra=bgra,
#     )

def _gen_single_data_lines_y_many_x(index_value, alpha=0.5):
    # Random Seed for reproducibility based on index
    np.random.seed(index_value * 555)

    # Generate multiple X arrays
    x_values = {}
    num_x_lines = 6
    num_points = 50 # Increased points for smoother sine wave
    for x_index in range(num_x_lines):
        # Generate x values relative to the index for variation
        x_values[f"x{x_index}"] = np.linspace(0, 4 * np.pi, num_points) + x_index * np.pi / 4

    # Calculate Y based on x0 as a shifted sine curve
    x0 = x_values["x0"]
    # Calculate phase shift based on index_value
    phase_shift = index_value * np.pi / 3
    # Calculate sine wave
    y_values = np.sin(x0 + phase_shift) + np.random.rand(num_points) * 0.2 # Add some noise

    # Determine color based on index
    color_name = COLORS[index_value % len(COLORS)]
    bgra_color = BGRA[color_name].copy() # Use copy to avoid modifying the original
    bgra_color[-1] = alpha # Set alpha

    return dict(
        y=y_values,
        **x_values,
        bgra=bgra_color,
    )

def _gen_single_data_lines_x_many_y(index_value, alpha=0.5):
    dd = _gen_single_data_lines_y_many_x(index_value, alpha=alpha)
    y_dict = {f"y{k[1:]}":v for k,v in dd.items() if k.startswith("x")}
    return dict(
        x=dd["y"],
        **y_dict,
        bgra=dd["bgra"],
    )

def _gen_single_data_contour(ii):
    # Random Seed
    np.random.seed(ii * 999)

    # Create grid data
    x = np.linspace(-5, 5, 10)
    y = np.linspace(-5, 5, 10)
    X, Y = np.meshgrid(x, y)

    # Create Z values (peaks with noise)
    sigma_x = 1.0 + 0.2 * ii
    sigma_y = 1.0 + 0.1 * ii

    # Create multiple peaks
    peaks = [
        (
            3,
            3,
            1 + 0.5 * ii,
            sigma_x,
            sigma_y,
        ),
        (-2, 2, 0.8 + 0.3 * ii, sigma_x * 0.8, sigma_y * 1.2),
        (0, -3, 0.9 + 0.4 * ii, sigma_x * 1.2, sigma_y * 0.8),
        (-3, -2, 0.7 + 0.2 * ii, sigma_x * 0.9, sigma_y * 0.9),
    ]

    Z = np.zeros_like(X)
    for x0, y0, height, sx, sy in peaks:
        Z += height * np.exp(
            -((X - x0) ** 2 / (2 * sx**2) + (Y - y0) ** 2 / (2 * sy**2))
        )

    # Add noise
    noise_level = 0.05 * (ii + 1)
    Z += np.random.normal(0, noise_level, Z.shape)

    # Convert to xyz format
    x_flat = X.flatten()
    y_flat = Y.flatten()
    z_flat = Z.flatten()

    return dict(
        x=x_flat,
        y=y_flat,
        z=z_flat,
        bgra=BGRA_FAKE,
    )


# Single
# ------------------------------
def _gen_single_data_bar(ii):
    # Random Seed
    np.random.seed(ii * 333)
    # X
    x = f"X {ii}"
    # Y
    y = 1.0 * (ii + 1) + np.random.normal(0, 0.3 * (ii + 1))
    yerr = 0.1 * (ii + 1)
    return dict(
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_barh(ii):
    # Random Seed
    np.random.seed(ii * 444)
    vv = _gen_single_data_bar(ii)
    return dict(
        y=vv["x"],
        x=vv["y"],
        xerr=vv["yerr"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_area(ii, alpha=0.5):
    # Random Seed
    np.random.seed(ii * 555)
    # X
    x = np.linspace(0, 10, 20) + ii
    # Y
    y = np.exp(-((x - 5 * (ii % 3)) ** 2) / 10)
    y += np.random.normal(0, 0.05 * (ii + 1), size=len(x))
    # y_lower = y - np.random.normal(0, 0.05 * (ii + 1), size=len(x))
    # y_upper = y + np.random.normal(0, 0.05 * (ii + 1), size=len(x))
    # Color
    bgra = BGRA[COLORS[ii % len(COLORS)]]
    bgra[-1] = alpha
    return dict(
        x=x,
        y=y,
        bgra=bgra,
    )


def _gen_single_data_box(ii):
    # Random Seed
    np.random.seed(ii * 666)
    # X
    x = f"Category {ii}"
    # Y
    # Generate data from uniform distribution to emphasize box plot visualization
    low = 3 * (ii + 1)
    high = 8 * (ii + 1)
    base_data = np.random.uniform(low, high, 30)
    # Add a few outliers outside the uniform range
    outliers_low = np.random.uniform(low - 2, low - 1, 1)
    outliers_high = np.random.uniform(high + 1, high + 2, 1)
    y = np.concatenate([base_data, outliers_low, outliers_high])
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_boxh(ii):
    vv = _gen_single_data_box(ii)
    return dict(
        y=vv["x"],
        x=vv["y"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )

def _gen_single_data_line(ii):
    # Random Seed
    np.random.seed(ii * 888)
    # X
    x = np.linspace(0, 10, 20)
    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )

def _gen_single_data_line_yerr(ii):
    # Random Seed
    np.random.seed(ii * 888)
    # X
    x = np.linspace(0, 10, 20)
    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    yerr = 0.2 * np.ones_like(x) * (1 + 0.1 * ii)
    return dict(
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_polar(ii):
    # Random Seed
    np.random.seed(ii * 123)

    # X
    theta = np.linspace(0, 2 * np.pi, 30)
    degree = theta / (2 * np.pi) * 360

    # Y
    r = 0.5 + ii + 0.5 * np.sin(theta * (ii + 1))
    r_fluctuation = np.random.normal(0, 0.1 * (ii + 1), size=len(theta))
    r = r + r_fluctuation
    return dict(
        theta=degree,
        r=r,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )

def _gen_single_data_scatter(ii):
    # Random Seed
    np.random.seed(ii * 789)
    n_points = 30 + ii * 5
    # X
    center_x = 5 * (ii % 3)
    x = np.random.normal(center_x, 1 + 0.2 * ii, n_points)
    # Y
    center_y = 5 * (ii // 3)
    y = np.random.normal(center_y, 1, n_points) + ii * 0.1 * x
    return dict(
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )

# EOF