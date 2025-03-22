#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-21 06:30:03 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_win.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/path/_to_win.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

def to_win(wsl_path):
    """
    Convert a WSL path to a Windows path.

    Args:
        wsl_path (str): Path in WSL format (e.g., /home/user/file.txt)

    Returns:
        str: Converted Windows path (e.g., C:\\Users\\user\\file.txt)
    """
    # Handle absolute paths
    if os.path.isabs(wsl_path):
        try:
            # Use wslpath command to convert the path
            result = subprocess.run(
                ['wslpath', '-w', wsl_path],
                capture_output=True,
                text=True,
                check=True
            )
            return result.stdout.strip()
        except (subprocess.SubprocessError, FileNotFoundError):
            # Fallback if wslpath doesn't work
            # Basic conversion for /mnt/c/ style paths
            if wsl_path.startswith('/mnt/'):
                drive = wsl_path[5:6].upper()
                path = wsl_path[7:].replace('/', '\\')
                return f"{drive}:{path}"
            return wsl_path
    # Handle relative paths by getting the absolute path first
    else:
        abs_path = os.path.abspath(wsl_path)
        return to_win(abs_path)

# EOF