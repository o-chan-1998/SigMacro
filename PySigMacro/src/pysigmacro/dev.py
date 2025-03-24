#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-24 19:50:47 (ywatanabe)"
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
DF = pd.DataFrame(
    columns=[ii for ii in range(30)], data=np.random.rand(100, 30)
)

# sp = ps.con.open(PATH)
# spw = ps.com.wrap(sp)
spw = ps.con.open(PATH)
notebooks = spw.Notebooks_obj
# print(notebooks.list)
idx = notebooks.find_indices(f"{PLOT_TYPE}")[0]
notebook = notebooks[idx]
# From here, templates defines indices and names
notebookitems = notebook.NotebookItems_obj
idx = notebookitems.find_indices(f"{PLOT_TYPE}_graph_S")[0]
graphitem = notebook.NotebookItems_obj[idx]
graphpages = graphitem.GraphPages_obj
graphpage = graphpages[0]
graphs = graphpage.Graphs_obj
graph = graphs[0]
graphitem.export_as_tif()


# # JPG Works
# spath = PATH.replace(".JNB", ".jpg")
# graphitem.Export(spath, "JPG")

# # TIFF does not work
# spath = PATH.replace(".JNB", ".tif")
# graphitem.Export(spath, "TIF")

# Or BMP (uncompressed, high quality but large file)
spath_bmp = PATH.replace(".JNB", ".bmp")
spath_cropped_bmp = PATH.replace(".JNB", "_cropped.bmp")
spath_tif = PATH.replace(".JNB", "_cropped.tif")

graphitem.Export(spath_bmp, "BMP")
ps.image.crop_images([spath_bmp])
ps.image.convert_bmp_to_tiff(spath_cropped_bmp)

# EOF