#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-23 12:50:41 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_import_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_import_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import time

import pandas as pd


def import_data(worksheet_item, df=None, csv=None):
    # df
    if (df is None) and (csv is not None):
        df = pd.read_csv(csv)

    # datatable object
    datatable_obj = worksheet_item.DataTable_obj

    # Header
    for ii, header in enumerate(df.columns):
        datatable_obj.ColumnTitle(ii, header)
        # time.sleep(0.1)

    # Data
    data_list = df.T.values.tolist()
    datatable_obj.PutData(data_list, 0, 0)

    # time.sleep(0.5)
    return datatable_obj

# EOF