#!/bin/bash
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-09 15:55:04 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/docs/git_add.sh

THIS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_PATH="$0.log"
touch "$LOG_PATH"


git add vba/ALL-IN-ONE-MACRO.vba
git add templates/jnb/*.JNB -f
git add templates/csv/*.csv -f
git add templates/gif/*.gif -f
git add templates/tif/*.tiff -f
git add templates/jnb/template.JNB -f
git add README.md

# EOF