<!-- ---
!-- Timestamp: 2025-04-09 15:57:07
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/README.md
!-- --- -->

# SigMacro

This package allows users to create publication-ready figures using [SigmaPlot](https://grafiti.com/sigmaplot-v16/) from Python, in a similar manner to matplotlib.

## Gallery
<div style="display: flex; flex-wrap: wrap; justify-content: space-between; max-width: 800px; margin: 0 auto;">
    <img src="templates/gif/line-line-line-line-line-line-line-line-line-line-line-line-line_cropped.gif" alt="Line Plot" width="150" />
    <img src="templates/gif/line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr-line_yerr_cropped.gif" alt="Line_Yerr Plot" width="150" />
    <img src="templates/gif/filled_line_cropped.gif" alt="Filled Line Plot" width="150" />
    <img src="templates/gif/area-area-area_cropped.gif" alt="Area Plot" width="150" />
    <img src="templates/gif/bar-bar-bar-bar-bar-bar-bar-bar-bar-bar-bar-bar-bar_cropped.gif" alt="Bar Plot" width="150" />
    <img src="templates/gif/barh-barh-barh-barh-barh-barh-barh-barh-barh-barh-barh-barh-barh_cropped.gif" alt="Horizontal Histogram Plot" width="150" />
    <img src="templates/gif/histogram-histogram-histogram_cropped.gif" alt="Histogram Plot" width="150" />
    <img src="templates/gif/scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter-scatter_cropped.gif" alt="Scatter Plot" width="150" />
    <img src="templates/gif/jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter-jitter_cropped.gif" alt="Jitter Plot" width="150" />
    <img src="templates/gif/box-box-box-box-box-box-box-box-box-box-box-box-box_cropped.gif" alt="Box Plot" width="150" />
    <img src="templates/gif/boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh-boxh_cropped.gif" alt="Horizontal Box Plot" width="150" />
    <img src="templates/gif/violin-violin-violin-violin-violin-violin-violin-violin-violin-violin-violin-violin-violin_cropped.gif" alt="Violin Plot" width="150" />
    <img src="templates/gif/contour_cropped.gif" alt="Contour Plot" width="150" />
    <img src="templates/gif/heatmap_cropped.gif" alt="Confusion Matrix" width="150" />
    <img src="templates/gif/polar-polar-polar-polar-polar-polar-polar-polar-polar-polar-polar-polar-polar_cropped.gif" alt="Polar Plot" width="150" />
</div>

## Working with GUI
<img src="./docs/demo.gif" alt="SigMacro Demo" width="400"/>

## How does it work?
#### In SigmaPlot:
1. [All-in-one macro](./vba/ALL-IN-ONE-MACRO.vba) is embedded in a [SigmaPlot template file](./templates/jnb/template.JNB) for:
   - Reading graph parameters
   - Plotting data

#### From Python:
1. Send plotting data and graph visualization parameters to SigmaPlot
2. Call SigmaPlot macros
3. Save (cropped) figures

In other words, [CSV files in these formats](./templates/csv) can be rendered by the [all-in-one-macro](./vba/ALL-IN-ONE-MACRO.vba). For more details, please refer to [the entry script](./PySigMacro/examples/demo.py) for the above demonstrations.

## Prerequisite
 - SigmaPlot License 
 - Windows OS

## Quick Start
```powershell
# Install pysigmacro package in development mode
cd PySigMacro
pip install -e .

# Run the demonstration script
python.exe ./PySigMacro/examples/demo.py
```

## TODO
- [ ] Implement simple interface like below
```python
import pysigmacro as psm
import pandas as pd

df = pd.DataFrame(...)
plotter = psm.Plotter()
plotter.add("area", df[["x", "y"]])
plotter.add("line", df[["x", "y", "yerr"]])
plotter.add("box", df[["x"]])
plotter.add("scatter", df[["x", "y"]])
# plotter.add("boxh", df[["y"]])
```

## Contact
Yusuke Watanabe (ywatanabe@alumni.u-tokyo.ac.jp)

<!-- EOF -->