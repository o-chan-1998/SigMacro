#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-25 17:39:35 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/dev.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/dev.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Scratch for Development
"""

import numpy as np
import pandas as pd

import pysigmacro as ps

# PARAMS
PLOT_TYPE = "line"
CLOSE_OTHERS = True
PATH = ps.path.copy_template("line", rf"C:\Users\wyusu\Downloads")
spw = ps.con.open(PATH)
notebooks = spw.Notebooks_obj
# print(notebooks.list)
notebook = notebooks[notebooks.find_indices(f"{PLOT_TYPE}")[0]]

# From here, templates defines indices and names
notebookitems = notebook.NotebookItems_obj
graphitem_s = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
]
graphitem_s.rename_xy_labels("aaa", "bbb")
worksheetitem = notebookitems[notebookitems.find_indices(f"{PLOT_TYPE}_worksheet")[0]]

graphitem_m = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_M")[0]
]
graphitem_l = notebookitems[
    notebookitems.find_indices(f"{PLOT_TYPE}_graph_L")[0]
]

# Data
DF = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

# spw = ps.con.open(PATH)
# notebooks = spw.Notebooks_obj
# # print(notebooks.list)
# notebook = notebooks[notebooks.find_indices(f"{PLOT_TYPE}")[0]]

# # From here, templates defines indices and names
# notebookitems = notebookitems
# graphitem_s = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
# ]
# graphitem_m = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_M")[0]
# ]
# graphitem_l = notebookitems[
#     notebookitems.find_indices(f"{PLOT_TYPE}_graph_L")[0]
# ]

# ps.utils.run_macro(
#     graphitem_s, "RenameXYLabels_macro", xlabel="X Label 1", ylabel="Y Label 1"
# )
# ps.utils.run_macro(
#     graphitem_m, "RenameXYLabels_macro", xlabel="X Label", ylabel="Y Label"
# )
# ps.utils.run_macro(
#     graphitem_l, "RenameXYLabels_macro", xlabel="X Label 2", ylabel="Y Label 2"
# )

# EOF