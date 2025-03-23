#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 13:15:45 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/dev.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/dev.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Scratch for Development
"""

import numpy as np
import pandas as pd

import pysigmacro as ps

# - [ ] Title, xlabel, ylabel: 8 pt
# - [ ] Tick labels: 7 pt
# - [ ] Legend: 6 pt
# - [ ] Tick length: 0.8 mm
# - [ ] Tick width: 0.2 mm
# - [ ] Hide top x-axis and right y-axis
# - [ ] Custom color selection implementation (reference: [01.Blue.vba](https://github.com/ywatanabe1989/SigmaPlot-

# PARAMS
PLOT_TYPE = "line"
CLOSE_OTHERS = True
PATH = ps.path.copy_template("line", rf"C:\Users\wyusu\Downloads")
DF = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

# EOF