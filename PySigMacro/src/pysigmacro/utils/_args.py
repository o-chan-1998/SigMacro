#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-25 03:57:54 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_to_args.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/utils/_to_args.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

ARGS_PATH = r"C:\Users\wyusu\Documents\SigMacro\.args.txt"

def to_args(**kwargs):
    with open(ARGS_PATH, 'w') as f:
        for k, v in kwargs.items():
            f.write(f"{k}={v}\n")

def list_args():
    args = {}
    try:
        with open(ARGS_PATH, 'r') as f:
            for line in f:
                line = line.strip()
                if '=' in line:
                    k, v = line.split('=', 1)
                    args[k] = v
    except FileNotFoundError:
        pass
    return args

# EOF