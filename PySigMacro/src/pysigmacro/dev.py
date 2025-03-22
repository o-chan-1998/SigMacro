#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-22 12:25:00 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/dev.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/dev.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Main SigmaPlot automation class
"""
import subprocess
import time
import pysigmacro as ps
import pandas as pd
import numpy as np
import win32com.client
from win32com.client import VARIANT
import pythoncom


class SigmaPlotAutomation:
    def __init__(self, plot_type, do_close=False, df=None):
        self.plot_type = plot_type
        self.df = df
        self.do_close = do_close
        self.section_name = f"{plot_type}_section"
        self.worksheet_name = f"{plot_type}_worksheet"
        self.graph_name = f"{plot_type}_graph"
        self.filepath = (
            f"C:\\User\\{os.getlogin()}\\Documents\\{plot_type}.JNB"
        )

    def process_connection(self):
        # Connection
        if self.do_close:
            ps.connect.close_all()
        sp = ps.connect.open()
        self.spw = ps.com.wrap(sp, "SigmaPlot")
        return self.spw

    def process_application(self):
        # Application
        self.app = self.spw.Application_obj
        return self.app

    def get_active_document(self):
        self.active_document = self.app.ActiveDocument_obj
        return self.active_document

    def process_notebooks(self):
        # Notebooks
        self.notebooks = self.app.Notebooks_obj
        self.notebooks.clear()
        return self.notebooks

    def process_notebook(self):
        # Notebook (= .JNB file)
        self.notebook = self.notebooks.Add()
        self.notebook.SaveAs(self.filepath)
        return self.notebook

    def process_notebookitems(self):
        # NotebookItems
        self.notebookitems = self.notebook.NotebookItems_obj
        # Section Item
        section_item = self.notebookitems["Section 1"]
        section_item.Name = self.section_name
        # Worksheet Item
        section_item = self.notebookitems["Data 1"]
        section_item.Name = self.worksheet_name
        # Graph Item
        self.notebookitems.add_graph(self.graph_name)
        return self.notebookitems

    def import_data(self):
        # Get worksheet item and datatable
        self.worksheet_item = self.notebookitems[self.worksheet_name]
        self.datatable = self.worksheet_item.DataTable_obj

        # Header
        for ii, header in enumerate(self.df.columns):
            self.datatable.ColumnTitle(ii, header)

        # Data
        df_T = self.df.T
        data_list = df_T.values.tolist()
        self.datatable.PutData(data_list, 0, 0)
        return self.datatable

    def create_graph(self):
        self.graph_item = self.notebookitems[self.graph_name]
        self.graph_pages = self.graph_item.GraphPages_obj
        self.graph_page = self.graph_pages[0]
        # Additional graph creation functionality would go here

    def save_notebook(self):
        self.notebook.Save()

    def run(self):
        self.process_connection()
        self.process_application()
        self.process_notebooks()
        self.process_notebook()
        self.process_notebookitems()
        if self.df is not None:
            self.import_data()
        self.create_graph()
        self.get_active_document()
        self.save_notebook()
        return self.filepath


# Usage example
if __name__ == "__main__":
    import pysigmacro as ps

    PLOT_TYPES = {
        "scatter": "Simple Scatter",
        "line": "Simple Straight Line",
        "line_and_scatter": "Simple Straight Line and Scatter",
        "bar": "Simple Vertical Bar",
        "bar_h": "Simple Horizontal Bar",
        "area": "Simple Area",
        "polar": "Polar Lines",
        "box": "Vertical Box Plot",
        "box_h": "Horizontal Box Plot",
        "contour": "Contour",
    }

    df = pd.DataFrame(
        columns=[
            "AA",
            "BB",
            "CC",
            "DD",
            "EE",
            "FF",
            "GG",
            "HH",
            "II",
            "JJ",
            "KK",
            "LL",
            "MM",
            "NN",
            "OO",
            "PP",
            "QQ",
            "RR",
            "SS",
            "TT",
            "UU",
            "VV",
            "WW",
            "XX",
            "YY",
            "ZZ",
        ],
        data=np.random.rand(100, 26),
    )

    for plot_type in PLOT_TYPES.keys():
        automator = SigmaPlotAutomation(plot_type, do_close=True, df=df)
        automator.run()


# # --------------------
# # Stable (Working)
# # --------------------


# def process_connection(do_close):
#     # Connection
#     if do_close:
#         ps.connect.close_all()
#     sp = ps.connect.open()
#     spw = ps.com.wrap(sp, "SigmaPlot")
#     return spw


# def process_application(spw):
#     # Application
#     app = spw.Application_obj
#     return app


# def get_active_document(app):
#     active_document = app.ActiveDocument_obj
#     return active_document


# def process_notebooks(app):
#     # Notebooks
#     notebooks = app.Notebooks_obj
#     notebooks.clear()
#     return notebooks


# def process_notebook(notebooks, spath):
#     # Notebook (= .JNB file)
#     notebook = notebooks.Add()
#     notebook.SaveAs(spath)
#     return notebook


# def process_notebookitems(notebook, section_name, worksheet_name, graph_name):
#     # NotebookItems
#     notebookitems = notebook.NotebookItems_obj

#     # Section Item
#     section_item = notebookitems["Section 1"]
#     section_item.Name = section_name

#     # Worksheet Item
#     section_item = notebookitems["Data 1"]
#     section_item.Name = worksheet_name

#     # Graph Item
#     notebookitems.add_graph(graph_name)

#     return notebookitems


# def import_data(datatable_obj, df):
#     # CSV Data
#     header_list = [list(df.columns)]
#     df_T = df.T
#     data_list = df_T.values.tolist()
#     datatable_obj.PutData(data_list, 0, 0)
#     return datatable_obj


# def main(spath, plot_type, df, do_close=False):
#     section_name = f"{PLOT_TYPE}_section"
#     worksheet_name = f"{PLOT_TYPE}_worksheet"
#     graph_name = f"{PLOT_TYPE}_graph"

#     spw = process_connection(do_close)
#     app = process_application(spw)
#     notebooks = process_notebooks(app)
#     notebook = process_notebook(notebooks, spath)
#     notebookitems = process_notebookitems(
#         notebook, section_name, worksheet_name, graph_name
#     )
#     worksheet_item = notebookitems[worksheet_name]
#     datatable = worksheet_item.DataTable_obj
#     datatable = import_data(datatable, df)
#     graph_item = notebookitems[graph_name]
#     graph_pages = graph_item.GraphPages_obj
#     active_document = get_active_document(app)
#     graph_page = graph_pages[0]
#     notebook.Save()


# # Params
# PLOT_TYPE = "bar"
# PATH = f"C:\\User\\wyusu\\Documents\\{PLOT_TYPE}.JNB"
# df = pd.DataFrame(columns=["aaa", "bbb"], data=np.random.rand(3, 2))
# DO_CLOSE = True
# main(PATH, PLOT_TYPE, df, DO_CLOSE)

# EOF