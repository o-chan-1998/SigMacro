<!-- ---
!-- Timestamp: 2025-03-22 12:38:20
!-- Author: ywatanabe
!-- File: /home/ywatanabe/win/documents/SigMacro/README.md
!-- --- -->

# SigMacro

A package for automating SigmaPlot routines.

![SigMacro Demo](./docs/demo.gif)

## Prerequisite

 - SigmaPlot License 
 - Windows OS

## Insallation

- SigmaPlot
  - A proprietary software for professional plotting (https://grafiti.com/sigmaplot-detail/)
  - Installation
    - [`./SigmaPlot_Installer/README.md`](./SigmaPlot_Installer/README.md)

- SigMacro
  - Series of macros for automating SigmaPlot
  - Installation
    - [`./SigMacro/README.md`](./SigMacro/README.md)

- PySigMacro
  - Python Interface for calling SigMacro
  - Installation
    - [`./PySigMacro/README.md`](./PySigMacro/README.md)

<details>
<summary>SigmaPlot Objects</summary>

``` plaintext
**Application**
└── **Notebooks** (collection)
    └── **Notebook**
        └── **NotebookItems** (collection)
            ├── **NativeWorksheetItem**
            │   ├── **DataTableNamedDataRanges** (collection)
            │   │   └── **NamedDataRange**
            │   ├── Smoother
            │   ├── PlotEquation
            │   └── **GraphWizard**
            ├── ExcelItem
            │   ├── DataTableNamedDataRanges (collection)
            │   │   └── NamedDataRange
            │   ├── Smoother
            │   ├── PlotEquation
            │   └── **GraphWizard**
            ├── FitItem
            │   └── FitResults
            ├── TransformItem
            ├── ReportItem
            ├── **MacroItem**
            ├── **NotebookItem**
            ├── **SectionItem**
            └── **GraphItem**
                └── **Pages** (collection)
                    └── **GraphObjects (Page)** (collection)
                        ├── Text
                        ├── **Line**
                        ├── **Solid**
                        ├── **GraphObject**
                        ├── Group
                        ├── Smoother
                        ├── PlotEquation
                        └── **Graph**
                            ├── **Graph Objects (Axis)** (collection)
                            │   └── **Axis**
                            ├── **Line** (collection)
                            ├── Text (collection)
                            │   └── Text
                            ├── Group (AutoLegend)
                            │   ├── Solid
                            │   └── Text
                            ├── **Graph Objects (Plots)** (collection)
                            │   └── **Plot**
                            │       ├── Symbol
                            │       ├── **Line**
                            │       ├── **Solid**
                            │       └── Text
                            ├── GraphObjects (Tuple) (collection)
                            │   └── Tuple
                            ├── Graph Objects (DropLines) (collection)
                            │   └── Line
                            └── Graph Objects (Function) (collection)
                                ├── Function (Line)
                                └── Text
```
</details>

## TODO
- [ ] Create a new graph (with graph_item.CreateWizardGraph())
- [ ] Add Plot
- [ ] Change Color
- [ ] Change Fig Size
- [ ] Ticks
  - [ ] Length
  - [ ] Width
  - [ ] Label
- [ ] X/Y Labels
- [ ] Title
- [ ] Export
- [ ] ConfusionMatrix

## Contact
Yusuke Watanabe (ywatanabe@alumni.u-tokyo.ac.jp)

<!-- EOF -->