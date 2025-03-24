#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-25 05:25:23 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import re

from ._BaseCOMWrapper import BaseCOMWrapper
from ._NotebookWrapper import NotebookWrapper
from ._NotebooksWrapper import NotebooksWrapper
from ._NotebookItemsWrapper import NotebookItemsWrapper
from ._WorksheetItemWrapper import WorksheetItemWrapper
from ._MacroItemWrapper import MacroItemWrapper
from ._GraphItemWrapper import GraphItemWrapper
from ._GraphPagesWrapper import GraphPagesWrapper
from ._GraphPageWrapper import GraphPageWrapper
from ._GraphsWrapper import GraphsWrapper
from ._GraphWrapper import GraphWrapper
from ._PlotsWrapper import PlotsWrapper
from ._base import register_wrap_function


def wrap(com_object, access_path="", path=""):
    """
    Wrap a COM object in an appropriate wrapper class.
    """
    try:
        # Ensure access_path is a string
        access_path = str(access_path) if access_path is not None else ""

        # Create and configure the appropriate wrapper
        wrapper = _create_wrapper(com_object, access_path)

        # Set path if provided
        if path:
            wrapper._path = path

        return wrapper
    except Exception as e:
        print(f"Error wrapping object: {e}")
        # Fall back to base wrapper
        wrapper = BaseCOMWrapper(com_object, access_path)
        if path:
            wrapper._path = path
        return wrapper


def _create_wrapper(com_object, access_path):
    """Create the appropriate wrapper based on access_path pattern"""
    access_path_last = access_path.split(".")[-1]
    # Notebooks
    if re.search(r"Notebooks$", access_path_last):
        return NotebooksWrapper(com_object, access_path)

    # Notebook
    elif re.search(r"Notebooks\[.*\]$", access_path_last):
        return NotebookWrapper(com_object, access_path)

    # NotebookItems
    elif re.search(r"NotebookItems$", access_path_last):
        return NotebookItemsWrapper(com_object, access_path)

    # Item
    elif re.search(r"NotebookItems\[.*\]$", access_path_last):

        # GraphItem
        if hasattr(com_object, "Name") and re.search(
            r".*_graph_.*", com_object.Name
        ):
            return GraphItemWrapper(com_object, access_path)
        # WorksheetItem
        elif hasattr(com_object, "Name") and re.search(
            r".*_worksheet_.*", com_object.Name
        ):
            return WorksheetItemWrapper(com_object, access_path)
        # WorksheetItem
        elif hasattr(com_object, "Name") and re.search(
            r".*_macro$", com_object.Name
        ):
            return MacroItemWrapper(com_object, access_path)
        else:
            return BaseCOMWrapper(com_object, access_path)

    # GraphPages
    elif re.search(r"GraphPages$", access_path_last):
        return GraphPagesWrapper(com_object, access_path)

    # GraphPage
    elif re.search(r"GraphPages\[.*\]$", access_path_last):
        return GraphPageWrapper(com_object, access_path)

    # Graphs
    elif re.search(r"Graphs$", access_path_last):
        return GraphsWrapper(com_object, access_path)

    # Graph
    elif re.search(r"Graphs\[.*\]$", access_path_last):
        return GraphWrapper(com_object, access_path)

    # Plots
    elif re.search(r"Plots$", access_path_last):
        return PlotsWrapper(com_object, access_path)

    # SigmaPlot root application
    elif access_path == "SigmaPlot":
        return BaseCOMWrapper(com_object, access_path)

    else:
        # Default to the base COM wrapper for unknown types
        return BaseCOMWrapper(com_object, access_path)


# def wrap(com_object, access_path=""):
#     """
#     Wrap a COM object in an appropriate wrapper class.
#     """
#     try:
#         # Notebooks
#         if re.search(r'Notebooks$', access_path):
#             return NotebooksWrapper(com_object, access_path)
#         # Notebook
#         elif re.search(r'Notebooks\[.*]$', access_path):
#             return NotebookWrapper(com_object, access_path)
#         # NotebookItems
#         elif re.search(r'NotebookItems$', access_path):
#             return NotebookItemsWrapper(com_object, access_path)
#         # Item
#         elif re.search(r'NotebookItems\[.*]$', access_path):
#             # GraphItem
#             if re.search(r'.*_graph_.*', com_object.Name):
#                 return GraphItemWrapper(com_object, access_path)
#             # WorksheetItem
#             elif re.search(r'.*_worksheet_.*', com_object.Name):
#                 return WorksheetWrapper(com_object, access_path)
#             else:
#                 return BaseCOMWrapper(com_object, access_path)
#         # GraphPages
#         elif re.search(r'GraphPages$', access_path):
#             return GraphPagesWrapper(com_object, access_path)
#         # GraphPage
#         elif re.search(r'GraphPages\[.*]$', access_path):
#             return GraphPageWrapper(com_object, access_path)
#         # Graphs
#         elif re.search(r'Graphs$', access_path):
#             return GraphsWrapper(com_object, access_path)
#         # Graph
#         elif re.search(r'Graps\[.*]$', access_path):
#             return GraphWrapper(com_object, access_path)
#         # Plots
#         elif re.search(r'Plots$', access_path):
#             return PlotsWrapper(com_object, access_path)
#         else:
#             # Default to the base COM wrapper for unknown types
#             return BaseCOMWrapper(com_object, access_path)
#     except Exception as e:
#         print(f"Error wrapping object: {e}")
#         # Fall back to base wrapper
#         return BaseCOMWrapper(com_object, access_path)

# def wrap(com_object, access_path=""):
#     """
#     Wrap a COM object in an appropriate wrapper class.
#     """
#     try:
#         # Select the appropriate wrapper class based on the path
#         if access_path.endswith("Notebooks[\d+]"):
#             return NotebooksWrapper(com_object, access_path)
#         if access_path.endswith("Notebooks"):
#             return NotebooksWrapper(com_object, access_path)
#         elif access_path.endswith("NotebookItems"):
#             return NotebookItemsWrapper(com_object, access_path)
#         elif access_path.endswith("GraphPages"):
#             return GraphPagesWrapper(com_object, access_path)
#         elif access_path.endswith("GraphPage"):
#             return GraphPageWrapper(com_object, access_path)
#         elif access_path.endswith("Graphs"):
#             return GraphsWrapper(com_object, access_path)
#         elif access_path.endswith("GraphItem"):
#             return GraphItemWrapper(com_object, access_path)
#         elif access_path.endswith("Plots"):
#             return PlotsWrapper(com_object, access_path)
#         else:
#             # Default to the base COM wrapper for unknown types
#             return BaseCOMWrapper(com_object, access_path)
#     except Exception as e:
#         print(f"Error wrapping object: {e}")
#         # Fall back to base wrapper
#         return BaseCOMWrapper(com_object, access_path)

# Register the wrap function
register_wrap_function(wrap)

# #!/usr/bin/env python3
# # -*- coding: utf-8 -*-
# # Timestamp: "2025-03-21 09:44:57 (ywatanabe)"
# # File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py
# # ----------------------------------------
# import os
# __FILE__ = (
#     "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/com/_wrap.py"
# )
# __DIR__ = os.path.dirname(__FILE__)
# # ----------------------------------------

# from ._BaseCOMWrapper import BaseCOMWrapper
# from ._NotebooksWrapper import NotebooksWrapper
# from ._NotebookItemsWrapper import NotebookItemsWrapper
# from ._base import register_wrap_function

# def wrap(com_object, path=""):
#     """Factory function to create appropriate wrapper"""
#     if path.endswith("Notebooks"):
#         return NotebooksWrapper(com_object, path)
#     if path.endswith("NotebookItems"):
#         return NotebookItemsWrapper(com_object, path)
#     else:
#         return BaseCOMWrapper(com_object, path)

# # Register the wrap function
# register_wrap_function(wrap)

# # EOF

# EOF