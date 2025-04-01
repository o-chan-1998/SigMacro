#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 16:29:00 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_csv.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/demo/_gen_csv.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ..path._to_win import to_win
from ..data._create_padded_df import create_padded_df
from ..data._create_graph_wizard_params import create_graph_wizard_params
from ._gen_data import gen_data
from ._gen_visual_params import gen_visual_params
from ._update_visual_params_with_nice_ticks import update_visual_params_with_nice_ticks
import numpy as np
import pandas as pd

# Demo data generation
# --------------------------------------------------


def gen_csv(plot_type, save=False):
    """
    Generate demo data for a given plot type and return as a DataFrame.

    Parameters:
        plot_type (str): The type of plot for which to generate data.
        save (bool): If True, the generated DataFrame will be saved as a CSV file.

    Returns:
        pandas.DataFrame: The generated demo data.
    """

    # Graph Wizard Parameters
    df_graph_wizard_params = create_graph_wizard_params(plot_type)

    # Parameters
    df_visual_params = gen_visual_params(plot_type)

    # Data
    df_data = gen_data(plot_type)

    # Update when auto specified
    df_visual_params = update_visual_params_with_nice_ticks(plot_type, df_visual_params, df_data)

    # Concatenate
    df = create_padded_df(df_graph_wizard_params, df_visual_params, df_data)

    if save:
        # Saving
        templates_dir = os.getenv(
            "SIGMACRO_TEMPLATES_DIR", os.path.join(__DIR__, "templates")
        )
        templates_csv_dir = os.path.join(templates_dir, "csv")
        spath = to_win(os.path.join(templates_csv_dir, f"{plot_type}.csv"))

        if os.path.exists(spath):
            os.remove(spath)

        df.to_csv(spath, index=False)
        print(f"Saved to: {spath}")

    return df

# EOF