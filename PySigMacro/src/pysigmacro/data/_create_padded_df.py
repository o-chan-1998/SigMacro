#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 19:58:46 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_padded_df.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_padded_df.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np
import pandas as pd


def create_padded_df(permutable_dict, filler=np.nan):
    """
    Create a padded pandas DataFrame from a dictionary with varying length items.

    This function takes a dictionary with keys as column names and values as
    column data, and converts it to a DataFrame. Values of different lengths
    are padded with a filler value to ensure rectangular data structure.

    Args:
        permutable_dict (dict): Dictionary where keys will be column names and
                              values will be column data in the resulting DataFrame.
        filler (any, optional): Value used to pad shorter lists. Defaults to np.nan.

    Returns:
        pandas.DataFrame: A DataFrame where each key in the input dictionary becomes
                        a column name, and all columns have the same length, with
                        shorter ones padded with the filler value.

    Notes:
        - Single values (str, int, float) are treated as length 1 lists
        - If a list of pandas Series is provided, they'll be converted to dictionaries
        - Original dictionary is copied to prevent modification
    """

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


def is_listed_X(obj, types):
    """
    Example:
        obj = [3, 2, 1, 5]
        _is_listed_X(obj,
    """
    import numpy as np

    try:
        condition_list = isinstance(obj, list)

        if not (isinstance(types, list) or isinstance(types, tuple)):
            types = [types]

        _conditions_susp = []
        for typ in types:
            _conditions_susp.append(
                (np.array([isinstance(o, typ) for o in obj]) == True).all()
            )

        condition_susp = np.any(_conditions_susp)

        _is_listed_X = np.all([condition_list, condition_susp])
        return _is_listed_X

    except:
        return False

# EOF