Option Explicit
' ========================================
' General Constants
' ========================================
Const DEBUG_MODE As Boolean = False
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"

' ========================================
' Axis Constants
' ========================================
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const HORIZONTAL As Long = 0
Const VERTICAL As Long = 1
Const MAJOR_TICK_INDEX As Long = 2

' ========================================
' Visual Settings
' ========================================
' Display Options
Const HIDE_LEGEND As Long = 0
Const HIDE_TITLE As Long = 0
' Styling
Const TICK_THICKNESS_INVISIBLE As Variant = &H00000000
Const TICK_THICKNESS_00008 As Variant = &H00000008
Const TICK_LENGTH_00032 As Variant = &H00000020
Const SSA_COLOR_ALPHA As Long = &H000008a7&
' Font Sizes
Const LABEL_PTS_07 As Long = 97
Const LABEL_PTS_08 As Long = 111
' Line Thickness
Const POLAR_LINE_THICKNESS As Double = 0.008 * 1000  ' Convert inches to internal units (1000 = 1 inch)

' ----------------------------------------
' Now, example formats of plotting data changed as follows. Note that indices are changeable and not all plot types will be included in a notebook (jnb file). Also, one type of plot may appear multiple times in a notebook (e.g., multiple line plots):
' Area Plot
'gw_param_keys 0', 'gw_param_values 0', 'label', 'x', 'y', 'bgra',
' Vertical Bar Plot
'gw_param_keys 1', 'gw_param_values 1', 'label', 'x', 'y', 'yerr', 'bgra',
' Horizontal Bar Plot
'gw_param_keys 2', 'gw_param_values 2', 'label', 'x', 'xerr', 'y', 'bgra',
' Vertical Box Plot
'gw_param_keys 3', 'gw_param_values 3', 'label', 'x', 'y', 'bgra',
' Horizontal Box Plot
'gw_param_keys 4', 'gw_param_values 4', 'label', 'y', 'x', 'bgra',
' Line Plot
'gw_param_keys 5', 'gw_param_values 5', 'label', 'x', 'y', 'yerr', 'bgra',
' Polar Plot
'gw_param_keys 6', 'gw_param_values 6', 'label', 'theta', 'r', 'bgra',
' Scatter Plot
'gw_param_keys 7', 'gw_param_values 7', 'label', 'x', 'y', 'bgra',
' Vertical Violin Plot
'gw_param_keys 8', 'gw_param_values 8', 'label', 'x_lower', 'x', 'x_upper', 'y', 'bgra',
' Horizontal Violin Plot
'gw_param_keys 9', 'gw_param_values 9', 'label', 'y_lower', 'y', 'y_upper', 'x', 'bgra',
' Filled Line Plot
'gw_param_keys 10', 'gw_param_values 10', 'label', 'x', 'y_lower', 'y', 'y_upper', 'bgra',
' Contour Plot
'gw_param_keys 11', 'gw_param_values 11', 'label', 'x', 'y', 'z',
' Confusion Matrix Plot
'gw_param_keys 12', 'gw_param_values 12', 'label', 'x', 'y', 'z', 'class_names'],
' ----------------------------------------

' ----------------------------------------
' Thus, we should revise this SigmaPlot macro script to perform:
' 1. Find start and end of a chunk of iPlot
' 2. Read the Graph Wizard Params
' 3. Determine the plot type
' 4. Get the data for plotting for the chunk
' 5. (Create graph if not exists and) add the plot
' ----------------------------------------

' ----------------------------------------
' Additionally, we need to check the type of data are consistent
' category, linear, Log, ...
' ----------------------------------------

' ----------------------------------------
' Also, we need to assume
' One contour occupies one figure (and so that one JNB file)
' This is the same for confusion matrix and filled line
' ----------------------------------------


' Rows
Const LABEL_ROW As Long = -1
' Graph Wizard-related
Const GW_PLOT_TYPE_ROW As Long = 0
Const GW_PLOT_STYLE_ROW As Long = 1
Const GW_DATA_TYPE_ROW As Long = 2
Const _GW_COLUMNS_PER_GW_ROW As Long = 3
Const _GW_GW_COLUMNS_COUNT_ARRAY_ROW As Long = 4
Const GW_DATA_SOURCE_ROW As Long = 5
Const GW_POLARUNITS_ROW As Long = 6
Const GW_ANGLEUNITS_ROW As Long = 7
Const GW_MIN_ANGLE_ROW As Long = 8
Const GW_MAX_ANGLE_ROW As Long = 9
Const GW_UNKONWN1_ROW As Long = 10
Const GW_GROUP_STYLE_ROW As Long = 11
Const GW_USE_AUTOMATIC_LEGENDS_ROW As Long = 12

' Graph Parameters
' ----------------------------------------
' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 0
Const GRAPH_PARAMS_COL As Long = 1
' Rows
Const X_LABEL_ROW As Long = 0
Const X_LABEL_ROTATION_ROW As Long = 1
Const X_MM_ROW As Long = 2
Const X_SCALE_TYPE_ROW As Long = 3
Const X_MIN_ROW As Long = 4
Const X_MAX_ROW As Long = 5
Const Y_LABEL_ROW As Long = 6
Const Y_LABEL_ROTATION_ROW As Long = 7
Const Y_MM_ROW As Long = 8
Const Y_SCALE_TYPE_ROW As Long = 9
Const Y_MIN_ROW As Long = 10
Const Y_MAX_ROW As Long = 11
' Ticks (Not handled by macros but embedded in JNB file)
' ----------------------------------------
Const X_TICKS_COL As Long = 2
Const Y_TICKS_COL As Long = 3
' Data Columns
' ----------------------------------------
' General
Const MAX_NUM_PLOTS As Long = 13
Const GW_START_COL_NAME_BASE As String = "gw_param_keys "
Const GW_START_COL As Long = -1 ' Will be found
' Const GW_ID_LABEL_EXPLANATION As Long = 0
' Const GW_ID_GW_PARAMS_COL As Long = 1

Const GW_ID_RGBA As Long = -1  ' This will be calculated dynamically in _CalculateColorColumnForPlot

' Add missing constants for column indexes within each chunk
Const GW_ID_PARAM_KEYS As Long = 0
Const GW_ID_PARAM_VALUES As Long = 1
Const GW_ID_LABEL As Long = 2

' Colors
Const RGB_BLACK As Long = &H00000000


' ========================================
' Helper
' ========================================
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Sub DebugType(item)
    If DEBUG_MODE Then
        MsgBox "Type: " & TypeName(item)
    End If
End Sub

Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object
    Dim cellValue As Variant
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function

Private Function ConvMmToInch(ByVal dValMm As Double) As Double
    ConvMmToInch = dValMm / 0.0254
End Function

Private Function ConvPtToInch(ByVal dValPt As Double) As Double
    ConvPtToInch = dValPt * (1000 / 72)
End Function

' ========================================
' Finder
' ========================================
Function _GetMaxCol() As Long
    Dim maxCol As Long, maxRow As Long, dataTable As Object
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    DataTable.GetMaxUsedSize(maxCol, maxRow)
    _GetMaxCol = maxCol
End Function

Function _FindColIdx(columnName As String) As Long
    Dim maxCol As Long, ColIndex As Long, ColName As String, ii As Long
    maxCol = _GetMaxCol()
    ColIndex = -1
    For ii = 0 To maxCol
        ColName = _ReadCell(ii, LABEL_ROW)
        If LCase(ColName) = LCase(columnName) Then
            ColIndex = ii
            Exit For
        End If
    Next ii
    _FindColIdx = ColIndex
End Function

Function _FindChunkStartCol(iPlot As Long) As Long
    Dim colName As String
    colName = GW_START_COL_NAME_BASE & iPlot
    _FindChunkStartCol = _FindColIdx(colName)
End Function

Function _FindChunkEndCol(iPlot As Long) As Long
    Dim startCol As Long, nextStartCol As Long
    Dim maxCol As Long
    
    startCol = _FindChunkStartCol(iPlot)
    
    If startCol = -1 Then
        _FindChunkEndCol = -1
        Exit Function
    End If
    
    nextStartCol = _FindChunkStartCol(iPlot + 1)
    
    If nextStartCol = -1 Then
        maxCol = _GetMaxCol()
        _FindChunkEndCol = maxCol
    Else
        _FindChunkEndCol = nextStartCol - 1
    End If
End Function

' ========================================
' Plot
' ========================================
' Revise Plot function to handle different plot types
Sub Plot()
' Open the worksheet
ActiveDocument.NotebookItems(WORKSHEET_NAME).Open

' Loop through all plot types
Dim iPlot As Long
Dim graphAlreadyExists As Boolean
graphAlreadyExists = _CheckGraphExists()

For iPlot = 0 To MAX_NUM_PLOTS - 1
    ' Find the start and end columns for this plot type
    Dim startCol As Long, endCol As Long
    startCol = _FindChunkStartCol(iPlot)
    
    ' If no more plot chunks found, exit loop
    If startCol = -1 Then
        Exit For
    End If
    
    endCol = _FindChunkEndCol(iPlot)
    DebugMsg "Plot " & iPlot & " columns: " & startCol & " to " & endCol
    
    ' Read GW parameters for this plot
    Dim plotType As String, plotStyle As String, dataType As String
    Dim dataSource As String, polarUnits As String, angleUnits As String
    Dim minAngle As Double, maxAngle As Double, groupStyle As String
    Dim useAutomaticLegends As Boolean, unknown1 As Variant
    
    ' Read parameters from the param_keys and param_values columns
    Dim valuesCol As Long
    valuesCol = startCol + 1  ' gw_param_values follows gw_param_keys
    
    ' Get type and style based on plot index
    plotType = _ReadCell(valuesCol, GW_PLOT_TYPE_ROW)    
    plotStyle = _ReadCell(valuesCol, GW_PLOT_STYLE_ROW)
    dataType = _ReadCell(valuesCol, GW_DATA_TYPE_ROW)
    dataSource = _ReadCell(valuesCol, GW_DATA_SOURCE_ROW)
    polarUnits = _ReadCell(valuesCol, GW_POLARUNITS_ROW)
    angleUnits = _ReadCell(valuesCol, GW_ANGLEUNITS_ROW)
    minAngle = CDbl(_ReadCell(valuesCol, GW_MIN_ANGLE_ROW))
    maxAngle = CDbl(_ReadCell(valuesCol, GW_MAX_ANGLE_ROW))
    unknown1 = _ReadCell(valuesCol, GW_UNKONWN1_ROW)
    groupStyle = _ReadCell(valuesCol, GW_GROUP_STYLE_ROW)
    useAutomaticLegends = CBool(_ReadCell(valuesCol, GW_USE_AUTOMATIC_LEGENDS_ROW))
    
    DebugMsg "Plot type: " & plotType & ", Style: " & plotStyle
    
    ' Build column mapping based on the plot type
    Dim ColumnsPerPlot() As Variant
    ColumnsPerPlot = _GetColumnMapping(iPlot, plotType, plotStyle, startCol, endCol)
    
    ' Get the column count array
    Dim PlotColumnCountArray() As Variant
    PlotColumnCountArray = _GetPlotCountColumnArray(plotType)
    
    ' Create the plot
    If Not graphAlreadyExists And iPlot = 0 Then
        DebugMsg "Creating new graph..."
        ' First plot with no existing graph - create the graph
        ActiveDocument.CurrentPageItem.CreateWizardGraph(plotType, _
            plotStyle, _
            dataType, _
            ColumnsPerPlot, _
            PlotColumnCountArray, _
            dataSource, _
            polarUnits, _
            angleUnits, _
            minAngle, _
            maxAngle, _
            unknown1, _
            groupStyle, _
            useAutomaticLegends)
        DebugMsg "New graph created"
        graphAlreadyExists = True
    Else
        ' If graph exists and this isn't a special plot type (contour/confusion matrix)
        If Not _IsSpecialPlotType(iPlot) Then
            ActiveDocument.NotebookItems(GRAPH_NAME).Open
            ActiveDocument.CurrentPageItem.AddWizardPlot(plotType, _
                plotStyle, _
                dataType, _
                ColumnsPerPlot, _
                PlotColumnCountArray, _
                dataSource, _
                polarUnits, _
                angleUnits, _
                minAngle, _
                maxAngle, _
                unknown1, _
                groupStyle, _
                useAutomaticLegends)
            DebugMsg "Plot added to existing graph"
        End If
    End If
Next iPlot

' Open the graph 
ActiveDocument.NotebookItems(GRAPH_NAME).Open
End Sub

' Check if this is a special plot type that needs its own graph
Function _IsSpecialPlotType(plotIndex As Long) As Boolean
    _IsSpecialPlotType = (plotIndex = 11 Or plotIndex = 12 Or plotIndex = 10)
End Function

' Get column mapping based on plot type and style
Function _GetColumnMapping(iPlot As Long, plotType As String, plotStyle As String, startCol As Long, endCol As Long) As Variant
    Dim mapping()
    Dim offset As Long
    
    ' Label is always at offset 2
    offset = 2
    
    Select Case iPlot
        Case 0  ' Area Plot: label, x, y, bgra
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            
        Case 1  ' Vertical Bar Plot: label, x, y, yerr, bgra
            ReDim mapping(2, 2)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            mapping(0, 2) = startCol + offset + 3  ' yerr
            
        Case 2  ' Horizontal Bar Plot: label, x, xerr, y, bgra
            ReDim mapping(2, 2)
            mapping(0, 0) = startCol + offset + 3  ' y
            mapping(0, 1) = startCol + offset + 1  ' x
            mapping(0, 2) = startCol + offset + 2  ' xerr
            
        Case 3  ' Vertical Box Plot: label, x, y, bgra
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            
        Case 4  ' Horizontal Box Plot: label, y, x, bgra
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1  ' y
            mapping(0, 1) = startCol + offset + 2  ' x
            
        Case 5  ' Line Plot: label, x, y, yerr, bgra
            ReDim mapping(2, 2)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            mapping(0, 2) = startCol + offset + 3  ' yerr
            
        Case 6  ' Polar Plot: label, theta, r, bgra
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1  ' theta
            mapping(0, 1) = startCol + offset + 2  ' r
            
        Case 7  ' Scatter Plot: label, x, y, bgra
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            
        Case 8  ' Vertical Violin: label, x_lower, x, x_upper, y, bgra
            ReDim mapping(2, 3)
            mapping(0, 0) = startCol + offset + 2  ' x
            mapping(0, 1) = startCol + offset + 4  ' y
            mapping(0, 2) = startCol + offset + 1  ' x_lower 
            mapping(0, 3) = startCol + offset + 3  ' x_upper
            
        Case 9  ' Horizontal Violin: label, y_lower, y, y_upper, x, bgra
            ReDim mapping(2, 3)
            mapping(0, 0) = startCol + offset + 4  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            mapping(0, 2) = startCol + offset + 1  ' y_lower
            mapping(0, 3) = startCol + offset + 3  ' y_upper
            
        Case 10  ' Filled Line: label, x, y_lower, y, y_upper, bgra
            ReDim mapping(2, 3)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 3  ' y
            mapping(0, 2) = startCol + offset + 2  ' y_lower
            mapping(0, 3) = startCol + offset + 4  ' y_upper
            
        Case 11  ' Contour Plot: label, x, y, z
            ReDim mapping(2, 2)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            mapping(0, 2) = startCol + offset + 3  ' z
            
        Case 12  ' Confusion Matrix: label, x, y, z, class_names
            ReDim mapping(2, 3)
            mapping(0, 0) = startCol + offset + 1  ' x
            mapping(0, 1) = startCol + offset + 2  ' y
            mapping(0, 2) = startCol + offset + 3  ' z
            mapping(0, 3) = startCol + offset + 4  ' class_names
            
        Case Else
            ' Default to a basic XY plot
            ReDim mapping(2, 1)
            mapping(0, 0) = startCol + offset + 1
            mapping(0, 1) = startCol + offset + 2
    End Select
    
    ' Fill in the row ranges for all columns
    Dim i As Integer
    For i = 0 To UBound(mapping, 2)
        mapping(1, i) = 0
        mapping(2, i) = 31999999
    Next i
    
    _GetColumnMapping = mapping
End Function

' Calculate the color column index
' Fixed _CalculateColorColumnForPlot function to use _FindColIdx for finding bgra column

Function _CalculateColorColumnForPlot(iPlot As Long) As Long
Dim startCol As Long, endCol As Long
Dim ii As Long, colName As String
startCol = _FindChunkStartCol(iPlot)
If startCol = -1 Then
_CalculateColorColumnForPlot = -1
Exit Function
End If
endCol = _FindChunkEndCol(iPlot)
' Look for a column named "bgra" within this chunk
For ii = startCol To endCol
colName = _ReadCell(ii, LABEL_ROW)
If LCase(colName) = "bgra" Then
_CalculateColorColumnForPlot = ii
Exit Function
End If
Next ii
' If no "bgra" column found, use the offset method as a fallback
_CalculateColorColumnForPlot = endCol - 1
End Function

Function _GetPlotCountColumnArray(plotType As String) As Variant
    Dim countArray()
    ReDim countArray(0)
    Select Case plotType
        Case "Vertical Bar Chart"
            countArray(0) = 3
        Case "Horizontal Bar Chart"
            countArray(0) = 3
        Case "Line Plot" 
            countArray(0) = 3
        Case "Scatter Plot"
            countArray(0) = 2
        Case "Filled Line Plot"
            countArray(0) = 4
        Case "Area Plot"
            countArray(0) = 2
        Case "Box Plot"
            countArray(0) = 2
        Case "Violin Plot"
            countArray(0) = 4
        Case "Polar Plot"
            countArray(0) = 2
        Case "Contour Plot"
            countArray(0) = 3
        Case "Confusion Matrix"
            countArray(0) = 4
        Case Else
            countArray(0) = 3
    End Select
    _GetPlotCountColumnArray = countArray
End Function

Function _CheckGraphExists() As Boolean
    On Error Resume Next
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Dim tempGraph As Object
        Set tempGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not tempGraph Is Nothing Then
            _CheckGraphExists = True
            Exit Function
        End If
    End If
    ' DebugMsg "No graph found"
    _CheckGraphExists = False
End Function
' ========================================
' Removers
' ========================================
Function RemoveExistingGraphs() As Boolean
    On Error Resume Next
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
    ActiveDocument.CurrentItem.SelectAll
    ActiveDocument.CurrentItem.Clear
End Function
Sub RemoveLegend()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_AUTOLEGENDSHOW, HIDE_LEGEND)
End Sub
Sub RemoveTopSpine()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub
Sub RemoveRightSpine()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub
Sub RemoveTitle()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_SHOWNAME, 0)
End Sub
' ========================================
' Color Setters
' ========================================
Function _SelectPlotObject(plotIndex As Long) As Object
    On Error Resume Next
    Set _SelectPlotObject = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not _SelectPlotObject Is Nothing Then
        On Error Resume Next
        _SelectPlotObject.SetObjectCurrent
        If Err.Number <> 0 Then
            DebugMsg "Error setting plot " & plotIndex & " as current: " & Err.Description
            Err.Clear
        End If
    Else
        Dim plotCount As Long
        plotCount = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots.Count
    End If
End Function
Function _DetectPlotType() As String
    On Error GoTo ErrorHandler
    Dim ObjectType As Variant
    Dim object_type As Variant
    object_type = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(0).GetAttribute(SLA_TYPE, ObjectType)
    If object_type = False Then
        _DetectPlotType = "Error: Failed to get object type."
        Exit Function
    End If
    ' Mapping
    Select Case object_type
        Case SLA_TYPE_SCATTER, SLA_TYPE_MINVAL, SLA_TYPE_POLARXY, SLA_TYPE_3DBAR, SLA_TYPE_TERNARYSCATTER
            _DetectPlotType = "LINE/SCATTER"
        Case SLA_TYPE_BAR
            _DetectPlotType = "BAR"
        Case SLA_TYPE_STACKED
            _DetectPlotType = "STACKED BAR"
        Case SLA_TYPE_TUKEY
            _DetectPlotType = "BOX"
        Case SLA_TYPE_3DSCATTER
            _DetectPlotType = "3D SCATTER/LINE"
        Case SLA_TYPE_MESH
            _DetectPlotType = "MESH"
        Case SLA_TYPE_CONTOUR
            _DetectPlotType = "CONTOUR"
        Case SLA_TYPE_POLAR
            _DetectPlotType = "POLAR"
        Case SLA_TYPE_MAXVAL
            _DetectPlotType = "MAXVAL"
        Case Else
            _DetectPlotType = "UNKNOWN OBJECT TYPE: " & object_type
    End Select
    ' DebugMsg "Type Detected: " & _DetectPlotType
    Exit Function
ErrorHandler:
    _DetectPlotType = "An error has occurred: " & Err.Description
End Function
Sub _SelectGraphObject(plotIndex As Long)
    On Error Resume Next
    Dim plotObj As Object
    Set plotObj = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not plotObj Is Nothing Then
        plotObj.SetObjectCurrent
        If Err.Number <> 0 Then
            ' DebugMsg "Error in _SelectGraphObject: " & Err.Description
            Err.Clear
        End If
    Else
        DebugMsg "Plot object not found in _SelectGraphObject for index " & plotIndex
    End If
End Sub
Sub _ChangeColorLine(RGB_VAL As Long, plotIndex As Long)
    ' SEA = Set Line Attribute
    ' DebugMsg "_ChangeColorLine called"
    _SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, 2)
    End With
End Sub
Sub _ChangeColorSymbol(RGB_VAL As Long, ALPHA_VAL As Long, plotIndex As Long)
    ' SSA = Set Symbol Attribute
    ' DebugMsg "_ChangeColorSymbol called"
    _SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLORREPEAT, 2)
    End With
End Sub
Sub _ChangeColorSolid(RGB_VAL As Long, plotIndex As Long)
    ' SDA = Set Solid Attribute
    ' Solids include graph planes, bars, and drawn solids objects
    ' DebugMsg "_ChangeColorSolid called"
    _SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLORREPEAT, 2)
    End With
End Sub
Sub _ChangeColorEdgeBlack(plotIndex As Long)
    ' SDA = Set Solid Attribute
    ' DebugMsg "_ChangeColorEdgeBlack called"
    _SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub
Sub _ChangeColorErrorBar(RGB_VAL As Long, plotIndex As Long)
    ' SLA = Set Line Attributes
    _SelectGraphObject plotIndex
    ' DebugMsg "_ChangeColorErrorBar called"
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLORREPEAT, 2)
    End With
End Sub
Sub _ChangeColorBox(RGB_VAL As Long, plotIndex As Long)
    ' SDA = Set Solid Attribute
    ' Solids include graph planes, bars, and drawn solids objects
    ' DebugMsg "_ChangeColorBox called"
    _SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Function _GetRGBFromColumn(columnIndex As Long) As Long
    ' DebugMsg "_GetRGBFromColumn called for plot " & columnIndex
    Dim rValue As Variant, gValue As Variant, bValue As Variant
    ' Read RGB values from worksheet (R, G, B values are assumed to be in adjacent columns)
    bValue = _ReadCell(columnIndex, 0)
    gValue = _ReadCell(columnIndex, 1)
    rValue = _ReadCell(columnIndex, 2)
    ' Convert to integers and create RGB color
    Dim r As Integer, g As Integer, b As Integer
    b = CInt(bValue)
    g = CInt(gValue)
    r = CInt(rValue)
    ' Standard RGB (VBA default)
    _GetRGBFromColumn = RGB(r, g, b)
End Function
Function _GetAlphaFromColumn(columnIndex As Long) As Long
    Dim alphaValue As Variant
    alphaValue = _ReadCell(columnIndex, 3)
    _GetAlphaFromColumn = alphaValue
End Function
Sub SetColors()
    On Error GoTo ErrorHandler
    Dim plotCount As Long
    Dim iPlot As Long
    Dim colorColumn As Long
    Dim RGB_VAL As Long
    Dim ALPHA_VAL As Long
    Dim graphItem As Object
    Dim plotObj As Object
    Dim DetectedPlotType As String
    ' Find the type of the object
    DetectedPlotType = _DetectPlotType()
    ' Get the graph page
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If graphItem Is Nothing Then
        DebugMsg "Error: Graph object not found"
        Exit Sub
    End If
    ' Get the number of plots
    plotCount = graphItem.Plots.Count
    ' Loop through all plots
    For iPlot = 0 To plotCount - 1
        colorColumn = _CalculateColorColumnForPlot(iPlot)
        RGB_VAL = _GetRGBFromColumn(colorColumn)
        ALPHA_VAL = _GetAlphaFromColumn(colorColumn)
        ' Apply color based on plot type
        Select Case DetectedPlotType
            Case "LINE/SCATTER"
                _ChangeColorLine RGB_VAL, iPlot
                _ChangeColorSymbol RGB_VAL, ALPHA_VAL, iPlot
                _ChangeColorSolid RGB_VAL, iPlot
            Case "3DSCATTER"
                _ChangeColorLine RGB_VAL, iPlot
                _ChangeColorSymbol RGB_VAL, ALPHA_VAL, iPlot
            Case "BAR", "STACKED"
                _ChangeColorSolid RGB_VAL, iPlot
                _ChangeColorEdgeBlack iPlot
            Case "BOX"
                _ChangeColorBox RGB_VAL, iPlot
            Case "POLAR"
                _ChangeColorLine RGB_VAL, iPlot
            Case Else
                DebugMsg "Unknown plot type detected: " & DetectedPlotType
        End Select
    Next iPlot
    Exit Sub
ErrorHandler:
    DebugMsg "Error in Main: " & Err.Description
End Sub
' ========================================
' Figure Size
' ========================================
Function _cvtMmToInternalUnit(mm As Long)
    _cvtMmToInternalUnit = mm*30
End Function
Sub SetFigureSize()
    On Error Resume Next
    ' Width
    Dim xLength_mm As Double
    Dim xLength_sp As Double
    xLength_mm = _ReadCell(GRAPH_PARAMS_COL, X_MM_ROW)
    xLength_sp = _cvtMmToInternalUnit(xLength_mm)
    ' Height
    Dim yLength_mm As Double
    Dim yLength_sp As Double
    yLength_mm = _ReadCell(GRAPH_PARAMS_COL, Y_MM_ROW)
    yLength_sp = _cvtMmToInternalUnit(yLength_mm)
    With ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        .Width = xLength_sp
        .Height = yLength_sp
    End With
End Sub
' ========================================
' Label Size
' ========================================
Sub _SetXLabelSize()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub
Sub _SetYLabelSize()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub
Sub SetXYLabelSizes()
    _SetXLabelSize()
    _SetYLabelSize()
End Sub
' ========================================
' Label Text
' ========================================
Function _SetXLabel()
    Dim xLabel As Variant
    xLabel = _ReadCell(GRAPH_PARAMS_COL, X_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, xLabel)
End Function
Function _SetYLabel()
    Dim yLabel As Variant
    yLabel = _ReadCell(GRAPH_PARAMS_COL, Y_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, yLabel)
End Function
Sub SetXYLabels()
    _SetXLabel()
    _SetYLabel()
End Sub
' ========================================
' Range
' ========================================
Sub _SetXRange()
    Dim xMin As String
    Dim xMax As String
    Dim xScaleType As Variant
    Dim xAxis As Object
    Const USE_CONSTANT_VALUE As Integer = 10
    ' Get the scale type
    xScaleType = _ReadCell(GRAPH_PARAMS_COL, X_SCALE_TYPE_ROW)
    ' Skip range setting for category or datetime axes
    If LCase(CStr(xScaleType)) = "category" Or LCase(CStr(xScaleType)) = "7" Or _
       LCase(CStr(xScaleType)) = "datetime" Or LCase(CStr(xScaleType)) = "date" Or _
       LCase(CStr(xScaleType)) = "time" Or LCase(CStr(xScaleType)) = "8" Then
        Exit Sub
    End If
    xMin = _ReadCell(GRAPH_PARAMS_COL, X_MIN_ROW)
    xMax = _ReadCell(GRAPH_PARAMS_COL, X_MAX_ROW)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, USE_CONSTANT_VALUE)
    ' ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, xMin)
    ' ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, xMax)
    If xMin <> "None" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, xMin)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 42991617)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, 20972546)
    End If
    If xMax <> "None" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, xMax)
    End If
End Sub
Sub _SetYRange()
    Dim yMin As String
    Dim yMax As String
    Dim yScaleType As Variant
    Dim yAxis As Object
    Const USE_CONSTANT_VALUE As Integer = 10
    ' Get the scale type
    yScaleType = _ReadCell(GRAPH_PARAMS_COL, Y_SCALE_TYPE_ROW)
    ' Skip range setting for category or datetime axes
    If LCase(CStr(yScaleType)) = "category" Or LCase(CStr(yScaleType)) = "7" Or _
       LCase(CStr(yScaleType)) = "datetime" Or LCase(CStr(yScaleType)) = "date" Or _
       LCase(CStr(yScaleType)) = "time" Or LCase(CStr(yScaleType)) = "8" Then
        Exit Sub
    End If
    yMin = _ReadCell(GRAPH_PARAMS_COL, Y_MIN_ROW)
    yMax = _ReadCell(GRAPH_PARAMS_COL, Y_MAX_ROW)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_OPTIONS, USE_CONSTANT_VALUE)
    If yMin <> "None" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_FROMVAL, yMin)
    End If
    If yMax <> "None" Then
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTRSTRING, SAA_TOVAL, yMax)
    End If
End Sub
Sub SetRanges()
    _SetXRange()
    _SetYRange()
End Sub
' ========================================
' Scales
' ========================================
Function _SetAxisType(axisIndex As Long, scaleType As Long)
    Dim axis As Object
    ' Get the axis object directly
    Set axis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(axisIndex)
    ' Set scale type
    axis.SetAttribute(SAA_TYPE, scaleType)
End Function
Function _GetScaleType(cellValue As Variant) As Long
    Const SAA_TYPE_LINEAR = 1
    Const SAA_TYPE_COMMON = 2
    Const SAA_TYPE_LOG = 3
    Const SAA_TYPE_PROBABILITY = 4
    Const SAA_TYPE_PROBIT = 5
    Const SAA_TYPE_LOGIT = 6
    Const SAA_TYPE_CATEGORY = 7
    Const SAA_TYPE_DATETIME = 8

    Dim scaleType As Long
    ' Convert string or number to appropriate scale type constant
    Select Case CStr(LCase(cellValue))
        Case "linear", "1"
            scaleType = SAA_TYPE_LINEAR
        Case "common", "common log", "2"
            scaleType = SAA_TYPE_COMMON
        Case "log", "natural log", "3"
            scaleType = SAA_TYPE_LOG
        Case "probability", "4"
            scaleType = SAA_TYPE_PROBABILITY
        Case "probit", "5"
            scaleType = SAA_TYPE_PROBIT
        Case "logit", "6"
            scaleType = SAA_TYPE_LOGIT
        Case "category", "7"
            scaleType = SAA_TYPE_CATEGORY
        Case "datetime", "date", "time", "8"
            scaleType = SAA_TYPE_DATETIME
        Case Else
            ' Default to linear if unrecognized
            scaleType = SAA_TYPE_LINEAR
    End Select
    _GetScaleType = scaleType
End Function
Sub _SetXScale()
    On Error Resume Next
    Dim xScaleData As Variant
    Dim xScaleType As Long
    ' Read scale types from worksheet
    xScaleData = _ReadCell(GRAPH_PARAMS_COL, X_SCALE_TYPE_ROW)
    ' Convert to scale type constants
    xScaleType = _GetScaleType(xScaleData)
    ' Set X axis scale type
    _SetAxisType HORIZONTAL, xScaleType
    On Error GoTo 0
End Sub
Sub _SetYScale()
    On Error Resume Next
    Dim yScaleData As Variant
    Dim yScaleType As Long
    ' Read scale types from worksheet
    yScaleData = _ReadCell(GRAPH_PARAMS_COL, Y_SCALE_TYPE_ROW)
    ' Convert to scale type constants
    yScaleType = _GetScaleType(yScaleData)
    ' Set X axis scale type
    _SetAxisType VERTICAL, yScaleType
    On Error GoTo 0
End Sub
Sub SetScales()
    _SetXScale()
    _SetYScale()
End Sub
' ========================================
' Ticks
' ========================================
Sub _SetXTicks()
    Dim xTicksData As Variant
    xTicksData = _ReadCell(X_TICKS_COL, 0)
    If Not xTicksData = "None" Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, X_TICKS_COL)
    End If
End Sub
Sub _SetYTicks()
    Dim yTicksData As Variant
    yTicksData = _ReadCell(Y_TICKS_COL, 0)
    If Not yTicksData = "None" Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, Y_TICKS_COL)
    End If
End Sub
Sub SetTicks()
    _SetXTicks()
    _SetYTicks()
End Sub
' ========================================
' Tick Sizes
' ========================================
Sub _SetXTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
End Sub
Sub _SetYTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_00008)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_00032)
End Sub
Sub SetTickSizes()
    _SetXTickSizes()
    _SetYTickSizes()
End Sub
' ' ========================================
' ' Tick Label Sizes
' ' ========================================
' Sub _SetXTickLabelSizes()
'     ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
'     ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
'     ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, STA_SIZE, LABEL_PTS_07)
' End Sub
' Sub _SetYTickLabelSizes()
'     ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
'     ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
'     ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, STA_SIZE, LABEL_PTS_07)
' End Sub
' Sub SetTickLabelSizes()
'    _SetXTickLabelSizes()
'    _SetYTickLabelSizes()
' End Sub
Sub SetXYLabelSizesAndTickLabelSizes()
    Dim oGraph As Object
    Dim oAxisX As Object
    Dim oAxisY As Object
    Dim oTextX As Object
    Dim oTextY As Object
    Dim oTextXTick As Object
    Dim oTextYTick As Object
    Set oGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    Set oAxisX = oGraph.Axes(HORIZONTAL)
    Set oAxisY = oGraph.Axes(VERTICAL)
    Set oTextX = oAxisX.AxisTitles(0)
    Set oTextY = oAxisY.AxisTitles(0)
    Set oTextXTick = oAxisX.TickLabelAttributes(MAJOR_TICK_INDEX)
    Set oTextYTick = oAxisY.TickLabelAttributes(MAJOR_TICK_INDEX)
    oTextX.SetAttribute(STA_SELECT, -65536)
    oTextY.SetAttribute(STA_SELECT, -65536)
    oTextX.SetAttribute(STA_SIZE, LABEL_PTS_08)
    oTextY.SetAttribute(STA_SIZE, LABEL_PTS_08)
    oTextXTick.SetAttribute(STA_SIZE, LABEL_PTS_07)
    oTextYTick.SetAttribute(STA_SIZE, LABEL_PTS_07)
End Sub
Sub SetLineWidth()
    Dim plotCount As Long
    Dim iPlot As Long
    Dim graphItem As Object
    Dim detectedPlotType As String
    detectedPlotType = _DetectPlotType()
    ' Get the graph page
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If graphItem Is Nothing Then
        DebugMsg "Error: Graph object not found in SetLineWidth"
        Exit Sub
    End If
    ' Get the number of plots
    plotCount = graphItem.Plots.Count
    ' Loop through all plots and set line width
    For iPlot = 0 To plotCount - 1
        _SelectGraphObject iPlot
        ' For polar plots, set the specific line width
        If detectedPlotType = "POLAR" Then
            ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_THICKNESS, POLAR_LINE_THICKNESS)
        End If
    Next iPlot
End Sub
Sub SetTickLabelRotation()
    Dim xRotation As Long
    Dim yRotation As Long
    Dim oGraph As Object
    Dim oAxisX As Object
    Dim oAxisY As Object
    Dim oTextXTick As Object
    Dim oTextYTick As Object
    ' Default rotations (0 degrees)
    xRotation = 0
    yRotation = 0
    ' Try to read rotation values from worksheet if available
    On Error Resume Next
    ' Assuming rotation values might be stored in cells next to the axis properties
    xRotation = CLng(_ReadCell(GRAPH_PARAMS_COL, X_LABEL_ROTATION_ROW)) * 10
    yRotation = CLng(_ReadCell(GRAPH_PARAMS_COL, Y_LABEL_ROTATION_ROW)) * 10
    On Error GoTo 0
    ' Set the tick label rotation
    Set oGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    Set oAxisX = oGraph.Axes(HORIZONTAL)
    Set oAxisY = oGraph.Axes(VERTICAL)
    Set oTextXTick = oAxisX.TickLabelAttributes(MAJOR_TICK_INDEX)
    Set oTextYTick = oAxisY.TickLabelAttributes(MAJOR_TICK_INDEX)
    ' Apply rotation values
    oTextXTick.SetAttribute(STA_ORIENTATION, xRotation)
    oTextYTick.SetAttribute(STA_ORIENTATION, yRotation)
End Sub
' ========================================
' Main
' ========================================
Sub Main()
    ' ========================================
    ' Working
    ' ========================================
    ' Remove any existing graphs
    RemoveExistingGraphs()
    ' Data Plotting
    Plot()
    ' Removers
    RemoveLegend()
    RemoveTopSpine()
    RemoveRightSpine()
    RemoveTitle()
    ' Color
    SetColors()
    ' Axes
    SetScales()
    SetRanges()
    ' Size
    SetFigureSize()
    ' Ticks
    SetTicks()
    SetTickSizes()
    ' SetTickLabelSizes() ' Only X
    ' Labels
    SetXYLabels()
    SetXYLabelSizes()
    ' Ticks and Labels
    SetXYLabelSizesAndTickLabelSizes()
    ' Tick label rotation
    SetTickLabelRotation()
    ' Line Width
    SetLineWidth()
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
End Sub
