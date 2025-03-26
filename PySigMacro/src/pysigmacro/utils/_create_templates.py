#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-26 18:56:57 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_create_templates.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_create_templates.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Template Creation Class
"""

import time

import numpy as np
import pandas as pd

from ..con._open import open as ps_con_open
from ..data._import_data import import_data as ps_data_import_data

class TemplateCreator:
    def __init__(
        self, plot_type, template_dir=None, close_others=False, df=None
    ):
        self.plot_type = plot_type
        self.df = df
        self.close_others = close_others

        # Names
        self.section_name = f"{plot_type}_section"
        self.worksheet_name = f"{plot_type}_worksheet"
        self.graph_name = f"{plot_type}_graph"

        # Path
        if not template_dir:
            template_dir = os.getenv(SIGMACRO_TEMPLATES_DIR, f"C:\\User\\{os.getlogin()}\\Documents")
        self.filepath = os.path.join(template_dir, f"{plot_type}.JNB")

    def process_connection(self, lpath=None, close_others=False):
        self.spw = ps_con_open(lpath=lpath, close_others=close_others)

    def process_application(self):
        self.app = self.spw.Application_obj

    def process_notebooks(self):
        self.notebooks = self.app.Notebooks_obj
        self.notebooks.clear()

    def process_notebook(self):
        self.notebook = self.notebooks.Add()
        if not os.path.exists(self.filepath):
            self.notebook.SaveAs(self.filepath)

    def process_notebookitems(
        self,
    ):
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

        time.sleep(1)

    def import_data(self, df=None):
        self.worksheet_item = self.notebookitems[self.worksheet_name]
        self.datatable = ps_data_import_data(self.worksheet_item, df)
        # time.sleep(1)

    def save_notebook(self):
        self.notebook.Save()

    def run(self):
        self.process_connection()
        self.process_application()
        self.process_notebooks()
        self.process_notebook()
        self.process_notebookitems()
        self.import_data(df=self.df)
        self.save_notebook()
        return self.filepath


def create_templates():
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
        creator = TemplateCreator(plot_type, close_others=True, df=df)
        creator.run()


if __name__ == "__main__":
    create_templates()

# EOF