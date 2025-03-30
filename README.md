<!-- ---
!-- Timestamp: 2025-03-30 10:30:45
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/README.md
!-- --- -->

# SigMacro

A package for automating SigmaPlot routines.

![SigMacro Demo](./docs/demo.gif)

<div style="display: grid; grid-template-columns: repeat(2, 1fr); grid-gap: 10px;">
    <img src="./SigMacro/Templates/line_cropped.gif" alt="Line Plot" width="150" />
    <img src="./SigMacro/Templates/filled_line_cropped.gif" alt="Filled Line Plot" width="150" />
    <img src="./SigMacro/Templates/area_cropped.gif" alt="Area Plot" width="150" />
    <img src="./SigMacro/Templates/scatter_cropped.gif" alt="Scatter Plot" width="150" />
    <img src="./SigMacro/Templates/bar_cropped.gif" alt="Bar Plot" width="150" />
    <img src="./SigMacro/Templates/bar_h_cropped.gif" alt="Horizontal Bar Plot" width="150" />
    <img src="./SigMacro/Templates/box_cropped.gif" alt="Box Plot" width="150" />
    <img src="./SigMacro/Templates/box_h_cropped.gif" alt="Horizontal Box Plot" width="150" />
    <img src="./SigMacro/Templates/violin_cropped.gif" alt="Violin Plot" width="150" />
    <img src="./SigMacro/Templates/contour_cropped.gif" alt="Contour Plot" width="150" />
    <img src="./SigMacro/Templates/polar_cropped.gif" alt="Polar Plot" width="150" />
    <img src="./SigMacro/Templates/confusion_matrix_cropped.gif" alt="Confusion Matrix" width="150" />
</div>

## Prerequisite

 - SigmaPlot License 
 - Windows OS

## Insallation

- SigmaPlot
  - A proprietary software for professional plotting (https://grafiti.com/sigmaplot-detail/)
  - Installation
    - [`./docs/v12_Installer/README.md`](./docs/v12_Installer/README.md)

- SigMacro
  - Series of macros for automating SigmaPlot
  - Installation
    - [`./SigMacro/README.md`](./SigMacro/README.md)

- PySigMacro
  - Python Interface for calling SigMacro
  - [Demo Movie](https://onedrive.live.com/?qt=allmyphotos&photosData=%2Fshare%2F12F1169924695EF9%213150863%3Fithint%3Dvideo%26e%3DLnoc26&sw=bypassConfig&cid=12F1169924695EF9&id=12F1169924695EF9%213150863&authkey=%21AFE1u69Zha9Sois&v=photos)
  - Installation
    - [`./PySigMacro/README.md`](./PySigMacro/README.md)

## Key Directories

``` bash
./PySigMacro/examples
./PySigMacro/src/pysigmacro/data/temp
```

## Usage

``` powershell
python.exe ./PySigMacro/examples/create_demo_data.py
```

## TODO
- [ ] As a Service

## Contact
Yusuke Watanabe (ywatanabe@alumni.u-tokyo.ac.jp)

<!-- EOF -->