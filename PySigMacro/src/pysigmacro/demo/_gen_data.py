#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 15:57:17 (ywatanabe)"
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

# Main
# ------------------------------


def gen_data(plot_type, n_plots=len(COLORS)):
    gen_data_func = {
        # Multiple versions
        "scatter": _gen_data_scatter,
        "line": _gen_data_line,
        "area": _gen_data_area,
        "bar": _gen_data_bar,
        "barh": _gen_data_barh,
        "box": _gen_data_box,
        "boxh": _gen_data_boxh,
        "violin": _gen_data_violin,
        "polar": _gen_data_polar,
        # Special
        "filled_line": _gen_data_filled_line,
        "contour": _gen_data_contour,
        "conf_mat": _gen_data_conf_mat,
    }[plot_type]
    return gen_data_func(n_plots)


# Core
# ------------------------------

# Special
# ------------------------------


def _gen_data_filled_line(*args, **kwargs):
    pass


def _gen_data_contour(*args, **kwargs):
    pass


def _gen_data_conf_mat(*args, **kwargs):
    pass


# Multiple versions
# ------------------------------


def _gen_data_scatter(n_plots=len(COLORS)):
    return _gen_multiple_data_base("scatter", n_plots=n_plots)


def _gen_data_bar(n_plots=len(COLORS)):
    return _gen_multiple_data_base("bar", n_plots=n_plots)


def _gen_data_barh(n_plots=len(COLORS)):
    return _gen_multiple_data_base("barh", n_plots=n_plots)


def _gen_data_line(n_plots=len(COLORS)):
    return _gen_multiple_data_base("line", n_plots=n_plots)


def _gen_data_area(n_plots=len(COLORS)):
    return _gen_multiple_data_base("area", n_plots=n_plots)


def _gen_data_box(n_plots=len(COLORS)):
    return _gen_multiple_data_base("box", n_plots=n_plots)


def _gen_data_boxh(n_plots=len(COLORS)):
    return _gen_multiple_data_base("boxh", n_plots=n_plots)


def _gen_data_polar(n_plots=len(COLORS)):
    return _gen_multiple_data_base("polar", n_plots=n_plots)


def _gen_data_violin(n_plots=len(COLORS)):
    return _gen_multiple_data_base("violin", n_plots=n_plots)


# Single
# ------------------------------


def _gen_single_data_bar(ii):
    # Random Seed
    np.random.seed(ii * 333)

    # X
    x = f"X {ii}"
    xerr = None

    # Y
    y = 1.0 * (ii + 1) + np.random.normal(0, 0.3 * (ii + 1))
    yerr = 0.1 * (ii + 1)

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_barh(ii):
    # Random Seed
    np.random.seed(ii * 444)

    # X
    x = 1.0 * (ii + 1) + np.random.normal(0, 0.3 * (ii + 1))
    xerr = 0.1 * (ii + 1)

    # Y
    y = f"Y {ii}"
    yerr = None

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_area(ii, alpha=0.5):
    # Random Seed
    np.random.seed(ii * 555)

    # X
    x = np.linspace(0, 10, 20) + ii
    xerr = None

    # Y
    y = np.exp(-((x - 5 * (ii % 3)) ** 2) / 10)
    y_lower = y - np.random.normal(0, 0.05 * (ii + 1), size=len(x))
    y_upper = y
    yerr = None


    bgra = BGRA[COLORS[ii % len(COLORS)]]
    bgra[-1] = alpha

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
        y_lower_value=y_lower,
        y_upper_value=y_upper,
    )


def _gen_single_data_box(ii):
    # Random Seed
    np.random.seed(ii * 666)

    # X
    x = f"Category {ii}"
    xerr = None

    # Y
    base_data = np.random.normal(5 * (ii + 1), 1 + ii * 0.2, 30)
    outliers = np.random.normal(5 * (ii + 1) + 3 * ((-1) ** ii), 0.5, 2)
    y = np.concatenate([base_data, outliers])
    yerr = None

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_boxh(ii):
    # Random Seed
    np.random.seed(ii * 777)

    # X
    base_data = np.random.normal(5 * (ii + 1), 1 + ii * 0.2, 30)
    outliers = np.random.normal(5 * (ii + 1) + 3 * ((-1) ** ii), 0.5, 2)
    x = np.concatenate([base_data, outliers])
    xerr = None

    # Y
    y = f"Category {ii}"
    yerr = None

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_line(ii):
    # Random Seed
    np.random.seed(ii * 888)

    # X
    x = np.linspace(0, 10, 20)
    xerr = None

    # Y
    y = np.sin(x + ii * 0.5) * (ii + 1)
    y += np.random.normal(0, 0.1 * (ii + 1), size=len(x))
    yerr = 0.2 * np.ones_like(x) * (1 + 0.1 * ii)

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_polar(ii):
    # Random Seed
    np.random.seed(ii * 123)

    # X
    theta = np.linspace(0, 2 * np.pi, 30)
    x_degree = theta / (2 * np.pi) * 360
    xerr = None

    # Y
    r = 0.5 + ii + 0.5 * np.sin(theta * (ii + 1))
    r_fluctuation = np.random.normal(0, 0.1 * (ii + 1), size=len(theta))
    y = r + r_fluctuation
    yerr = None

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x_degree,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_scatter(ii):
    np.random.seed(ii * 789)

    n_points = 30 + ii * 5

    # X
    center_x = 5 * (ii % 3)
    x = np.random.normal(center_x, 1 + 0.2 * ii, n_points)
    xerr = None

    # Y
    center_y = 5 * (ii // 3)
    y = np.random.normal(center_y, 1, n_points) + ii * 0.1 * x
    yerr = None

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
    )


def _gen_single_data_violin(ii):
    # Violin plot data - create multimodal distributions
    np.random.seed(ii * 42)
    # Create base position
    x = f"Category {ii}"
    # Create multimodal distribution for interesting violins
    # Mix two or three normal distributions
    if ii % 3 == 0:
        # Bimodal
        dist1 = np.random.normal(ii * 2, 0.5, 15)
        dist2 = np.random.normal(ii * 2 + 3, 0.5, 15)
        y = np.concatenate([dist1, dist2])
    elif ii % 3 == 1:
        # Trimodal
        dist1 = np.random.normal(ii * 1.5, 0.4, 10)
        dist2 = np.random.normal(ii * 1.5 + 2, 0.3, 15)
        dist3 = np.random.normal(ii * 1.5 + 4, 0.5, 10)
        y = np.concatenate([dist1, dist2, dist3])
    else:
        # Skewed
        dist1 = np.random.normal(ii * 2, 0.8, 20)
        dist2 = np.random.normal(ii * 2 + 2, 0.3, 10)
        y = np.concatenate([dist1, dist2])
    # No point in having xerr for violin plots
    xerr = None
    yerr = None
    # Can provide quartile information for box plots within violins
    y_lower = np.percentile(y, 25)
    y_upper = np.percentile(y, 75)

    return _gen_single_data_base_wrapper(
        ii=ii,
        x=x,
        xerr=xerr,
        y=y,
        yerr=yerr,
        y_lower_value=y_lower,
        y_upper_value=y_upper,
    )


# Base
# ------------------------------


def _gen_multiple_data_base(plot_type, n_plots=len(COLORS)):
    gen_single_data_func = {
        "scatter": _gen_single_data_scatter,
        "line": _gen_single_data_line,
        "bar": _gen_single_data_bar,
        "barh": _gen_single_data_barh,
        "area": _gen_single_data_area,
        "box": _gen_single_data_box,
        "boxh": _gen_single_data_boxh,
        "polar": _gen_single_data_polar,
        "violin": _gen_single_data_violin,
    }[plot_type]
    out_dict = {}
    for ii in range(n_plots):
        out_dict.update(gen_single_data_func(ii))

    # To df
    out_df = create_padded_df(out_dict)
    return out_df


def _gen_single_data_base(
    ii=None,
    x_label=None,
    x_value=None,
    xerr_label=None,
    xerr_value=None,
    y_label=None,
    y_value=None,
    yerr_label=None,
    yerr_value=None,
    bgra_label=None,
    bgra_value=None,
    x_lower_label=None,
    x_lower_value=None,
    x_upper_label=None,
    x_upper_value=None,
    y_lower_label=None,
    y_lower_value=None,
    y_upper_label=None,
    y_upper_value=None,
):
    ii_space = f" {ii}" if ii is not None else f""

    # X
    x_label = x_label if x_label is not None else f"X{ii_space}"
    x_value = x_value if x_value is not None else np.nan
    xerr_label = xerr_label if xerr_label is not None else f"X Err.{ii_space}"
    xerr_value = xerr_value if xerr_value is not None else np.nan
    x_lower_label = (
        x_lower_label if x_lower_label is not None else f"X Lower{ii_space}"
    )
    x_lower_value = x_lower_value if x_lower_value is not None else np.nan
    x_upper_label = (
        x_upper_label if x_upper_label is not None else f"X Upper{ii_space}"
    )
    x_upper_value = x_upper_value if x_upper_value is not None else np.nan

    # Y
    y_label = y_label if y_label is not None else f"Y{ii_space}"
    y_value = y_value if y_value is not None else np.nan
    yerr_label = yerr_label if yerr_label is not None else f"Y Err.{ii_space}"
    yerr_value = yerr_value if yerr_value is not None else np.nan
    bgra_label = bgra_label if bgra_label is not None else f"BGRA{ii_space}"
    y_lower_label = (
        y_lower_label if y_lower_label is not None else f"Y Lower{ii_space}"
    )
    y_lower_value = y_lower_value if y_lower_value is not None else np.nan
    y_upper_label = (
        y_upper_label if y_upper_label is not None else f"Y Upper{ii_space}"
    )
    y_upper_value = y_upper_value if y_upper_value is not None else np.nan

    if bgra_value is None:
        if ii is not None:
            bgra_value = BGRA[COLORS[ii % len(COLORS)]]
        else:
            bgra_value = BGRA[COLORS["black"]]

    return {
        x_label: x_value,
        xerr_label: xerr_value,
        x_lower_label: x_lower_value,
        x_upper_label: x_upper_value,
        y_label: y_value,
        yerr_label: yerr_value,
        y_lower_label: y_lower_value,
        y_upper_label: y_upper_value,
        bgra_label: bgra_value,
    }


def _gen_single_data_base_wrapper(
    ii=None,
    x=None,
    xerr=None,
    x_lower_value=None,
    x_upper_value=None,
    y=None,
    yerr=None,
    y_lower_value=None,
    y_upper_value=None,
    bgra=None,
):
    return _gen_single_data_base(
        ii=ii,
        x_value=x,
        xerr_value=xerr,
        x_lower_value=x_lower_value,
        x_upper_value=x_upper_value,
        y_value=y,
        yerr_value=yerr,
        y_lower_value=y_lower_value,
        y_upper_value=y_upper_value,
        bgra_value=bgra,
    )


# def _gen_single_data_filled_line(ii):
#     # Random Seed
#     np.random.seed(ii * 42)

#     # X
#     x = np.linspace(0, 10, 20) + ii
#     xerr = None

#     # Y
#     y = np.sin(x + ii * 0.5) * (ii + 1)
#     yerr = None
#     y_lower = (
#         y
#         - 0.5 * (ii + 1)
#         + np.random.normal(0, 0.3, size=len(x)) * (ii + 1) * 0.2
#     )
#     y_upper = (
#         y
#         + 0.5 * (ii + 1)
#         + np.random.normal(0, 0.4, size=len(x)) * (ii + 1) * 0.3
#     )

#     return _gen_single_data_base_wrapper(
#         ii=ii,
#         x=x,
#         xerr=xerr,
#         y=y,
#         yerr=yerr,
#         y_lower_value=y_lower,
#         y_upper_value=y_upper,
#     )

# EOF