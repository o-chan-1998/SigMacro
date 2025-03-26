#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 12:23:55 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_force_df.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_force_df.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np
import pandas as pd


def force_df(permutable_dict, filler=np.nan):

    if is_listed_X(permutable_dict, pd.Series):
        permutable_dict = [sr.to_dict() for sr in permutable_dict]
    ## Deep copy
    permutable_dict = permutable_dict.copy()

    ## Get the lengths
    max_len = 0
    for k, v in permutable_dict.items():
        # Check if v is an iterable (but not string) or treat as single length otherwise
        if isinstance(v, (str, int, float)) or not hasattr(v, "__len__"):
            length = 1
        else:
            length = len(v)
        max_len = max(max_len, length)

    ## Replace with appropriately filled list
    for k, v in permutable_dict.items():
        if isinstance(v, (str, int, float)) or not hasattr(v, "__len__"):
            permutable_dict[k] = [v] + [filler] * (max_len - 1)
        else:
            permutable_dict[k] = list(v) + [filler] * (max_len - len(v))

    ## Puts them into a DataFrame
    out_df = pd.DataFrame(permutable_dict)

    return out_df

# EOF