#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-05 06:05:42 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np

from ..const import BGRA, COLORS
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params
from scipy import stats

import pandas as pd

# Main
# ------------------------------

def gen_data(plot_types):

    chunks_df = pd.DataFrame()
    for i_plot, plot_type in enumerate(plot_types):
        try:
            gen_single_data_func = {
                "scatter": _gen_single_data_scatter,
                "line": _gen_single_data_line,
                "lineh": _gen_single_data_lineh,
                "bar": _gen_single_data_bar,
                "barh": _gen_single_data_barh,
                "area": _gen_single_data_area,
                "box": _gen_single_data_box,
                "boxh": _gen_single_data_boxh,
                "polar": _gen_single_data_polar,
                "violin": _gen_single_data_violin,
                "violinh": _gen_single_data_violinh,
                "conf_mat": _gen_single_data_conf_mat,
                "filled_line": _gen_single_data_filled_line,
                "contour": _gen_single_data_contour,
            }[plot_type]
            single_plot_dict = gen_single_data_func(i_plot)
            single_plot_df = create_padded_df(single_plot_dict)
            gw_df = create_graph_wizard_params(plot_type)
            gw_df.columns = [f"{col} {i_plot}" for col in gw_df.columns]
            single_chunk = create_padded_df(gw_df, single_plot_df)
            chunks_df = create_padded_df(chunks_df, single_chunk)
        except Exception as e:
            print(plot_type)
            print(e)

    return chunks_df


# Special
# ------------------------------
def _gen_single_data_filled_line(ii, alpha=0.5):
    plot_type = "filled_line"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 555)
    # X
    x = np.linspace(0, 10, 20) + ii
    # Y
    y = np.exp(-((x - 5 * (ii % 3)) ** 2) / 10)
    y_lower = y - np.random.normal(0, 0.05 * (ii + 1), size=len(x))
    y_upper = y + np.random.normal(0, 0.05 * (ii + 1), size=len(x))

    # Color
    bgra = BGRA[COLORS[ii % len(COLORS)]]
    bgra[-1] = alpha
    return dict(
        label=label,
        x=x,
        y_lower=y_lower,
        y=y,
        y_upper=y_upper,
        bgra=bgra,
    )


def _gen_single_data_filled_lineh(ii):
    plot_type = "filled_lineh"
    # Label
    label = f"{plot_type} {ii}"

    vv = _gen_single_data_filled_line(ii)

    # Color
    bgra = BGRA[COLORS[ii % len(COLORS)]]

    return dict(
        label=label,
        y=vv["x"],
        x_lower=vv["y_lower"],
        x=vv["y"],
        x_upper=vv["y_upper"],
        bgra=bgra,
    )


def _gen_single_data_contour(ii):
    plot_type = "contour"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 999)

    # Create grid data
    x = np.linspace(-5, 5, 50)
    y = np.linspace(-5, 5, 50)
    X, Y = np.meshgrid(x, y)

    # Create Z values (peaks with noise)
    sigma_x = 1.0 + 0.2 * ii
    sigma_y = 1.0 + 0.1 * ii

    # Create multiple peaks
    peaks = [
        (3, 3, 1 + 0.5 * ii, sigma_x, sigma_y),  # x, y, height, sigma_x, sigma_y
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
        label=label,
        x=x_flat,
        y=y_flat,
        z=z_flat,
    )


def _gen_single_data_conf_mat(ii, n_classes=4):
    plot_type = "conf_mat"
    # Label
    label = f"{plot_type} {ii}"

    # Generate class names based on n_classes
    class_names = [f"Class {chr(65+i)}" for i in range(n_classes)]

    # Random seed for reproducibility
    np.random.seed(ii * 777)

    # Generate random confusion matrix
    conf_matrix = np.random.randint(8000, 23000, size=(n_classes, n_classes))

    # Make diagonal values higher (better accuracy)
    for i in range(n_classes):
        conf_matrix[i, i] = np.random.randint(15000, 25000)

    # Convert to x, y, z format for SigmaPlot
    x_coords = []
    y_coords = []
    z_values = []

    for i in range(n_classes):
        for j in range(n_classes):
            x_coords.append(i)  # Row (true label)
            y_coords.append(j)  # Column (predicted label)
            z_values.append(conf_matrix[i, j])  # Count/value

    return dict(
        label=label,
        x=x_coords,
        y=y_coords,
        z=z_values,
        class_names=class_names,
    )


# Single
# ------------------------------
def _gen_single_data_bar(ii):
    plot_type = "bar"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 333)
    # X
    x = f"X {ii}"
    # Y
    y = 1.0 * (ii + 1) + np.random.normal(0, 0.3 * (ii + 1))
    yerr = 0.1 * (ii + 1)
    return dict(
        label=label,
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_barh(ii):
    plot_type = "barh"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 444)
    vv = _gen_single_data_bar(ii)
    return dict(
        label=label,
        x=vv["y"],
        xerr=vv["yerr"],
        y=vv["x"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_area(ii, alpha=0.5):
    plot_type = "area"
    # Label
    label = f"{plot_type} {ii}"
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
        label=label,
        x=x,
        y=y,
        bgra=bgra,
    )


def _gen_single_data_box(ii):
    plot_type = "box"
    # Label
    label = f"{plot_type} {ii}"
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
        label=label,
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_boxh(ii):
    plot_type = "boxh"
    # Label
    label = f"{plot_type} {ii}"
    vv = _gen_single_data_box(ii)
    return dict(
        label=label,
        y=vv["x"],
        x=vv["y"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_line(ii):
    plot_type = "line"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 888)
    # X
    x = np.linspace(0, 10, 20)
    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    yerr = 0.2 * np.ones_like(x) * (1 + 0.1 * ii)
    return dict(
        label=label,
        x=x,
        y=y,
        yerr=yerr,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_lineh(ii):
    plot_type = "lineh"
    # Label
    label = f"{plot_type} {ii}"
    vv = _gen_single_data_line(ii)
    return dict(
        label=label,
        y=vv["x"],
        x=vv["y"],
        xerr=vv["yerr"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_polar(ii):
    plot_type = "polar"
    # Label
    label = f"{plot_type} {ii}"
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
        label=label,
        theta=theta,
        r=r,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_scatter(ii):
    plot_type = "scatter"
    # Label
    label = f"{plot_type} {ii}"
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
        label=label,
        x=x,
        y=y,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_violin(ii):
    plot_type = "violin"
    # Label
    label = f"{plot_type} {ii}"
    # Random Seed
    np.random.seed(ii * 666)
    # Y (original sample points)
    # Generate data from uniform distribution to emphasize box plot visualization
    low = 3 * (ii + 1)
    high = 8 * (ii + 1)
    base_data = np.random.uniform(low, high, 30)
    # Add a few outliers outside the uniform range
    outliers_low = np.random.uniform(low - 2, low - 1, 1)
    outliers_high = np.random.uniform(high + 1, high + 2, 1)
    y = np.concatenate([base_data, outliers_low, outliers_high])
    # X
    x = f"Category {ii}"
    # Calculate kernel density estimate for violin edges
    a = np.sort(y)
    sigma = np.std(a)
    m = len(a)
    # Compute Inter Quartile Range for bandwidth
    q1 = np.percentile(a, 25)
    q3 = np.percentile(a, 75)
    iqr = q3 - q1
    # Compute bandwidth (Silverman's rule of thumb)
    scaled_iqr = iqr / 1.34
    min_val = min(sigma, scaled_iqr) if scaled_iqr > 0 else sigma
    bandwidth = 0.9 * min_val / (m**0.2)
    # Define range for kernel density
    z_start = np.min(a) - 3.5 * bandwidth
    z_end = np.max(a) + 3.5 * bandwidth
    n = 256
    # Compute kernel density values
    z_values = np.linspace(z_start, z_end, n)
    kde = stats.gaussian_kde(a, bw_method=bandwidth / sigma)
    density = kde(z_values) * bandwidth
    # Scale density to match desired width (0.4 = typical box plot width)
    width_factor = 0.4
    max_density = np.max(density) if len(density) > 0 else 0.1
    scaled_density = density * (width_factor / max_density)
    # Create x_lower and x_upper arrays for violin edges
    position = ii + 1
    x_lower = position - scaled_density
    x_upper = position + scaled_density
    return dict(
        label=label,
        x_lower=x_lower,
        x=position,
        x_upper=x_upper,
        y=z_values,
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )


def _gen_single_data_violinh(ii):
    plot_type = "violinh"
    # Label
    label = f"{plot_type} {ii}"
    vv = _gen_single_data_violin(ii)
    return dict(
        label=label,
        y_lower=vv["x_lower"],
        y=vv["x"],
        y_upper=vv["x_upper"],
        x=vv["y"],
        bgra=BGRA[COLORS[ii % len(COLORS)]],
    )

# EOF