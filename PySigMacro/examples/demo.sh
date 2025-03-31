#!/bin/bash
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-31 19:59:32 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/examples/demo.sh

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$0.log"
touch "$LOG_PATH"


# Function to get relative path from current directory
to_relative() {
    # Get the absolute path
    local path="$1"
    # Get the current working directory
    local cwd="$(pwd)"
    # Use Python to calculate the relative path
    python3 -c "import os.path; print(os.path.relpath('$path', '$cwd'))"
}

THIS_DIR_REL="$(to_relative $THIS_DIR)"

powershell.exe python "$THIS_DIR_REL"/create_demo_csv.py
sleep 1
powershell.exe python "$THIS_DIR_REL"/create_demo_jnb.py

# EOF