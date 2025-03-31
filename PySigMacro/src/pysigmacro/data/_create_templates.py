#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-31 20:21:55 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_templates.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_templates.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

"""
Template Creation Class
"""

import time

import numpy as np
import pandas as pd

from ..con._close_all import close_all as ps_con_close_all
from ..con._open import open as ps_con_open
from ..data._import_data import import_data as ps_data_import_data

class TemplateCreator:
    def __init__(
            self, plot_type, templates_dir=None, df=None, from_demo_csv=True,
    ):
        self.plot_type = plot_type
        self.df = df
        self.from_demo_csv = from_demo_csv

        # Names
        self.section_name = f"section"
        self.worksheet_name = f"worksheet"
        self.graph_name = f"graph"

        # Path
        if not templates_dir:
            templates_dir = os.getenv("SIGMACRO_TEMPLATES_DIR", os.path.join(__DIR__, "templates"))

        self.path_template = os.path.join(templates_dir, f"template.JNB")
        self.path_demo_csv_path = os.path.join(templates_dir, f"{plot_type}.csv")
        self.path_save = os.path.join(templates_dir, f"{plot_type}.JNB")

        assert os.path.exists(self.path_template)
        assert os.path.exists(self.path_demo_csv_path)

    def copy_template(self):
        if os.path.exists(self.path_save):
            os.remove(self.path_save)

        max_trials = 10
        while os.path.exists(self.path_save):
            time.sleep(1)
            max_trials += 1
            if max_trials > 10:
                time.sleep(1)
                break


        from ..path._copy import copy
        copy(self.path_template, self.path_save)

        print(f"Saved to: {self.path_save}")

    def process_connection(self):
        self.spw = ps_con_open(lpath=self.path_save)

    def process_application(self):
        self.app = self.spw.Application_obj

    def process_notebooks(self):
        self.notebooks = self.app.Notebooks_obj
        self.notebooks.clean()

    def process_notebook(self):
        filename = os.path.basename(self.path_save)
        if filename in self.notebooks.list:
            self.notebook = self.notebooks[self.notebooks.find_indices(filename)[0]]
        else:
            self.notebook = self.notebooks.Add()
            self.notebook.SaveAs(self.path_save)

    def process_notebookitems(
        self,
    ):
        # NotebookItems
        self.notebookitems = self.notebook.NotebookItems_obj
        self.notebookitems.clean()

        self.sectionitem = self.notebookitems[self.notebookitems.find_indices("section")[0]]
        self.worksheetitem = self.notebookitems[self.notebookitems.find_indices("worksheet")[0]]
        self.graphitem = self.notebookitems[self.notebookitems.find_indices("graph")[0]]
        self.all_in_one_macro = self.notebookitems[self.notebookitems.find_indices("all-in-one-macro")[0]]

    def import_data(self, df=None):
        if self.from_demo_csv:
            csv = self.path_demo_csv_path
        self.datatable = ps_data_import_data(self.worksheetitem, df=self.df, csv=csv)

    def run_all_in_one_macro(self):
        self.worksheetitem.activate()
        self.graphitem.activate()
        time.sleep(1)
        self.all_in_one_macro.run()
        self.graphitem.activate()
        time.sleep(1)

    def save_notebook(self):
        self.notebook.Save()

    def run(self):
        ps_con_close_all()
        self.copy_template()
        self.process_connection()
        self.process_application()
        self.process_notebooks()
        self.process_notebook()
        self.process_notebookitems()
        self.import_data()
        self.run_all_in_one_macro()
        self.graphitem.export_as_tif()
        self.save_notebook()
        ps_con_close_all()
        return self.path_save

def create_templates(plot_type):
    creator = TemplateCreator(plot_type, from_demo_csv=True)
    creator.run()

if __name__ == "__main__":
    from ..const._PLOT_TYPES import PLOT_TYPES

    for plot_type in PLOT_TYPES:
        df = None
        create_templates(plot_type, df=df)

# EOF