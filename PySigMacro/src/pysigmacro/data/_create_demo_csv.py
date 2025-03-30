#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-30 20:06:21 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_demo_csv.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_demo_csv.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

from ._gen_graph_wizard_params import gen_graph_wizard_params
from ._gen_visual_params import gen_visual_params
from ._gen_demo_data import gen_demo_data

from ._create_padded_df import create_padded_df
from ..path._to_win import to_win
import pandas as pd
import numpy as np

# Demo data generation
# --------------------------------------------------

def create_demo_csv(plot_type, save=False):
    """
    Generate demo data for a given plot type and return as a DataFrame.

    Parameters:
        plot_type (str): The type of plot for which to generate data.
        save (bool): If True, the generated DataFrame will be saved as a CSV file.

    Returns:
        pandas.DataFrame: The generated demo data.
    """

    # Graph Wizard Parameters
    df_graph_wizard_params = gen_graph_wizard_params(plot_type)

    # Parameters
    df_visual_params = gen_visual_params(plot_type)

    # Data
    df_data = gen_demo_data(plot_type)

    # Concatenate
    df = create_padded_df(df_graph_wizard_params, df_visual_params, df_data)

    if save:
        # Saving
        sdir = os.path.join(__DIR__, "templates")
        spath = to_win(os.path.join(sdir, f"{plot_type}.csv"))
        df.to_csv(spath, index=False)
        print(f"Saved to: {spath}")

    return df

# EOF