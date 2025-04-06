Option Explicit

' ========================================
' General Constants
' ========================================
' Const DEBUG_MODE As Boolean = True
' Const DEBUG_MODE As Boolean = False
Const GLOBAL_DEBUG_MODE As Boolean = False
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
Const AREAFILLTYPE_VERTICAL As Long = 1
Const SSA_COLOR_ALPHA As Long = &H000008a7&

' Font Sizes
Const LABEL_PTS_07 As Long = 97
Const LABEL_PTS_08 As Long = 111

' Line Thickness
Const POLAR_LINE_THICKNESS As Double = 0.008 * 1000
Const AREA_LINE_THICKNESS As Variant = 0

' Rows
Const LABEL_ROW As Long = -1

' ========================================
' Graph Wizard-related constants
' ========================================
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

' ========================================
' Entire Graph
' ========================================
' ----------------------------------------
' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 0
Const GRAPH_PARAMS_COL As Long = 1
Const X_TICKS_COL As Long = 2
Const Y_TICKS_COL As Long = 3

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


' Data Columns
' ----------------------------------------
' General
Const GW_START_COL_BASE_NAME As String = "gw_param_keys "
Const GW_START_COL As Long = -1
Const GW_ID_PARAM_KEYS As Long = 0
Const GW_ID_PARAM_VALUES As Long = 1
Const GW_ID_LABEL As Long = 2
Const GW_ID_RGBA As Long = -1
' Colors
Const RGB_BLACK As Long = &H00000000
Const RGB_TRANSPARENT As Long = 0



' ========================================
' Helper
' ========================================
Sub DebugMsg(DEBUG_MODE As Boolean, msg As String)
    If GLOBAL_DEBUG_MODE Or DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Sub DebugType(DEBUG_MODE As Boolean, item)
    If GLOBAL_DEBUG_MODE Or DEBUG_MODE Then
       MsgBox "Type: " & TypeName(item)
    End If
End Sub

' Reader
' ----------------------------------------

Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object
    Dim cellValue As Variant
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function

' Color
' ----------------------------------------
Function _GetRGBFromColumn(columnIndex As Long) As Long
    Const DEBUG_MODE As Boolean = False
    ' DebugMsg(DEBUG_MODE, "_GetRGBFromColumn called for plot " & columnIndex
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
    Const DEBUG_MODE As Boolean = False
    Dim alphaValue As Variant
    alphaValue = _ReadCell(columnIndex, 3) * 100
    _GetAlphaFromColumn = alphaValue
End Function

Function _ReadPlotTypeAsStr(iPlot As Long) As String
    Const DEBUG_MODE As Boolean = False
    Dim startCol As Long, valuesCol As Long, labelCol As Long
    Dim plotType As String
    Dim spacePos As Long
    
    startCol = _FindChunkStartCol(iPlot)
    If startCol <> -1 Then
        labelCol = startCol + GW_ID_LABEL
        plotType = _ReadCell(labelCol, 0)
        
        ' Extract base type by removing any trailing numbers
        spacePos = InStr(plotType, " ")
        If spacePos > 0 Then
            plotType = Left(plotType, spacePos - 1)
        End If
        
        _ReadPlotTypeAsStr = plotType
    Else
        _ReadPlotTypeAsStr = "line"
    End If
End Function

' Selector
' ----------------------------------------
Sub _SelectGraphObject(plotIndex As Long)
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim plotObj As Object
    Set plotObj = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not plotObj Is Nothing Then
        plotObj.SetObjectCurrent
        If Err.Number <> 0 Then
            ' DebugMsg(DEBUG_MODE, "Error in _SelectGraphObject: " & Err.Description
            Err.Clear
        End If
    Else
        DebugMsg(DEBUG_MODE, "Plot object not found in _SelectGraphObject for index " & plotIndex)
    End If
End Sub


' Finder
' ----------------------------------------
Function _GetMaxCol() As Long
    Const DEBUG_MODE As Boolean = False
    Dim maxCol As Long, maxRow As Long, dataTable As Object
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    DataTable.GetMaxUsedSize(maxCol, maxRow)
    _GetMaxCol = maxCol
End Function

Function _FindColIdx(columnName As String) As Long
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
    Dim colName As String
    colName = GW_START_COL_BASE_NAME & iPlot
    _FindChunkStartCol = _FindColIdx(colName)
End Function

Function _FindChunkEndCol(iPlot As Long) As Long
    Const DEBUG_MODE As Boolean = False
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
        _FindChunkEndCol = maxCol - 1
    Else
        _FindChunkEndCol = nextStartCol - 1
    End If
End Function

Function _GetNumPlots() As Long
    Const DEBUG_MODE As Boolean = False
    Dim iCol As Long
    Dim count As Long
    Dim maxCol As Long
    maxCol = _GetMaxCol()
    count = 0
    For iCol = 0 To maxCol
        If _FindChunkStartCol(iCol) <> -1 Then
            count = count + 1
        Else
            ' No more chunks found, exit loop
            Exit For
        End If
    Next iCol
    _GetNumPlots = count
    DebugMsg(DEBUG_MODE, "Found " & count & " plot chunks")
End Function

' Function _IsSpecialPlotType(plotType As String) As Boolean
'     Const DEBUG_MODE As Boolean = False
'     ' Check if plot type is one of the special types
'     _IsSpecialPlotType = (plotType = "Confusion Matrix" Or plotType = "Filled Line" Or plotType = "Contour")
' End Function

' Graph Wizard
' ----------------------------------------
Function _GetColumnMapping(startCol As Long, endCol As Long) As Variant
    Const DEBUG_MODE As Boolean = False
    Dim mapping()

    ' Data Columns
    Dim nDataCols As Long
    Const nHeadCols As Long = 3
    Const nTailCols As Long = 1
    nDataCols = (endCol - startCol + 1) - (nHeadCols + nTailCols)

    ReDim mapping(2, nDataCols)

    Dim iCol As Long
    For iCol = 0 To nDataCols
        mapping(0, iCol) = startCol + nHeadCols + iCol
    Next iCol

    ' Fill in the row ranges for all columns
    Dim ii As Integer
    For ii = 0 To UBound(mapping, 2)
        mapping(1, ii) = 0
        mapping(2, ii) = 31999999
    Next ii

    _GetColumnMapping = mapping
End Function

Function _GetPlotCountColumnArray(startCol As Long, endCol As Long) As Variant
    Const DEBUG_MODE As Boolean = False

    Dim countArray()
    ReDim countArray(0)

    ' Data Columns
    Dim nDataCols As Long
    Const nHeadCols As Long = 3
    Const nTailCols As Long = 1
    nDataCols = (endCol - startCol + 1) - (nHeadCols + nTailCols)

    DebugMsg(DEBUG_MODE, "_GetPlotCountColumnArray called")
    DebugMsg(DEBUG_MODE, "startCol: " & startCol)
    DebugMsg(DEBUG_MODE, "endCol: " & endCol)
    DebugMsg(DEBUG_MODE, "nDataCols: " & nDataCols)

    ' ReDim countArray(0)
    countArray(0) = nDataCols

    _GetPlotCountColumnArray = countArray
End Function

Function _DoesGraphExist() As Boolean
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Dim tempGraph As Object
        Set tempGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not tempGraph Is Nothing Then
            _DoesGraphExist = True
            Exit Function
        End If
    End If
    ' DebugMsg(DEBUG_MODE, "No graph found")
    _DoesGraphExist = False
End Function


' ========================================
' Plot
' ========================================
Sub Plot()
    Const DEBUG_MODE As Boolean = False
    ' Open the worksheet
    ActiveDocument.NotebookItems(WORKSHEET_NAME).Open

    ' Get the number of plots
    Dim numPlots As Long
    numPlots = _GetNumPlots()
    
    ' Loop through all plot types
    Dim iPlot As Long
    Dim graphAlreadyExists As Boolean
    graphAlreadyExists = _DoesGraphExist()

    For iPlot = 0 To numPlots - 1

        ' Find the start and end columns for this plot type
        Dim startCol As Long, endCol As Long
        startCol = _FindChunkStartCol(iPlot)

        ' If no more plot chunks found, exit loop
        If startCol = -1 Then
            DebugMsg(DEBUG_MODE, "No plot chunks found")
            Exit For
        End If

        endCol = _FindChunkEndCol(iPlot)
        DebugMsg(DEBUG_MODE, "Plot " & iPlot & " columns: " & startCol & " to " & endCol)

        ' Read GW parameters for this plot
        Dim plotType As String, plotStyle As String, dataType As String
        Dim dataSource As String, polarUnits As String, angleUnits As String
        Dim minAngle As Double, maxAngle As Double, groupStyle As String
        Dim useAutomaticLegends As Boolean, unknown1 As Variant
        ' Read parameters from the param_keys and param_values columns
        Dim valuesCol As Long
        valuesCol = startCol + 1

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

        ' Build column mapping based on the plot type
        Dim ColumnsPerPlot() As Variant
        ColumnsPerPlot = _GetColumnMapping(startCol, endCol)

        ' Get the column count array
        Dim PlotColumnCountArray() As Variant
        PlotColumnCountArray = _GetPlotCountColumnArray(startCol, endCol)

        ' Create the plot
        If Not graphAlreadyExists And iPlot = 0 Then
            ' If Not graphAlreadyExists Then
            DebugMsg(DEBUG_MODE, "Creating new graph...")
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
                                                           , _
                                                           groupStyle, _
                                                           useAutomaticLegends)
            DebugMsg(DEBUG_MODE, "New graph created")
            graphAlreadyExists = True
        Else
            ' If graph exists and this isn't a special plot type (contour/confusion matrix)
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
                                                      , _
                                                      groupStyle, _
                                                      useAutomaticLegends)
            DebugMsg(DEBUG_MODE, "Plot added to existing graph")
        End If
    Next iPlot
    ' Open the graph
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
End Sub


' ========================================
' Removers
' ========================================
Sub RemoveExistingGraphs()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
    ActiveDocument.CurrentItem.SelectAll
    ActiveDocument.CurrentItem.Clear
End Sub

Sub RemoveLegend()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_AUTOLEGENDSHOW, HIDE_LEGEND)
End Sub

Sub RemoveTopSpine()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub

Sub RemoveRightSpine()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_INVISIBLE)
End Sub

Sub RemoveTitle()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_SHOWNAME, 0)
End Sub

' ========================================
' Color
' ========================================
Sub SetColors()
    Const DEBUG_MODE As Boolean = False
    On Error GoTo ErrorHandler
    Dim plotCount As Long
    Dim iPlot As Long
    Dim colorColumn As Long
    Dim RGB_VAL As Long
    Dim ALPHA_VAL As Long
    Dim graphItem As Object
    Dim plotObj As Object
    Dim plotType As String

    ' Get the graph page
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If graphItem Is Nothing Then
        DebugMsg(DEBUG_MODE, "Error: Graph object not found")
        Exit Sub
    End If

    ' Get the number of plots
    plotCount = graphItem.Plots.Count

    ' Loop through all plots
    For iPlot = 0 To plotCount - 1
        plotType = LCase(_ReadPlotTypeAsStr(iPlot))
        colorColumn = _FindChunkEndCol(iPlot)
        RGB_VAL = _GetRGBFromColumn(colorColumn)
        ALPHA_VAL = _GetAlphaFromColumn(colorColumn)

        ' Apply color based on plot type
        Select Case plotType
           Case "area"
              _ApplyColorArea(iPlot, RGB_VAL, ALPHA_VAL)
           Case "bar", "barh"
              _ApplyColorBar(iPlot, RGB_VAL, ALPHA_VAL)
           Case "box", "boxh"
              _ApplyColorBox(iPlot, RGB_VAL, ALPHA_VAL)
           Case "line"
              _ApplyColorLine(iPlot, RGB_VAL, ALPHA_VAL)
           Case "polar"
              _ApplyColorPolar(iPlot, RGB_VAL, ALPHA_VAL)
           Case "scatter"
              _ApplyColorScatter(iPlot, RGB_VAL, ALPHA_VAL)
           Case "violin", "violinh"
              _ApplyColorViolin(iPlot, RGB_VAL, ALPHA_VAL)
           Case "filled_Line"
              _ApplyColorFilledLine(iPlot, RGB_VAL, ALPHA_VAL)
           Case "contour"
              _ApplyColorContour(iPlot, RGB_VAL, ALPHA_VAL)
           Case "conf_mat"
              _ApplyColorConfMat(iPlot, RGB_VAL, ALPHA_VAL)
           Case "3dscatter"
              _ApplyColor3DScatter(iPlot, RGB_VAL, ALPHA_VAL)
            Case Else
                DebugMsg(DEBUG_MODE, "Unknown plot type detected: " & plotType)
        End Select
    Next iPlot
    Exit Sub

ErrorHandler:
    DebugMsg(DEBUG_MODE, "Error in Main: " & Err.Description)
End Sub


Sub _ApplyColorArea(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorArea called")
    DebugMsg(DEBUG_MODE, "RGB_VAL: " & RGB_VAL & ", ALPHA_VAL: " & ALPHA_VAL)
    
    ' Select the plot object
    _SelectGraphObject(iPlot)

    ' koko
    
    ' Apply line color
    With ActiveDocument.CurrentPageItem
        ' Make line invisible by setting to transparent
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Fill area with color
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        
        ' Apply fill type and transparency
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_AREAFILLTYPE, AREAFILLTYPE_VERTICAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)

    End With
End Sub

Sub _ApplyColorBar(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorBar called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply bar colors
    With ActiveDocument.CurrentPageItem
        ' Solid attributes (for bar fill)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        
        ' Edge attributes (black edge)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Sub _ApplyColorBox(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorBox called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply box colors
    With ActiveDocument.CurrentPageItem
        ' Fill color
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        
        ' Edge color (black)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Sub _ApplyColorLine(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorLine called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply line color
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
        
        ' Solid attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        
        ' Error bar attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorPolar(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorPolar called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply line color only for polar plots
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorScatter(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorScatter called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply scatter colors
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
        
        ' Solid attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        
        ' Error bar attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorViolin(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorViolin called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply violin colors
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
        
        ' Solid attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        
        ' Error bar attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorFilledLine(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorFilledLine called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply filled line colors
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
        
        ' Solid attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        
        ' Error bar attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_ERRCOLOR, RGB_VAL)
    End With
End Sub

Sub _ApplyColorContour(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorContour called")
    ' Custom handling for contour plots if needed
End Sub

Sub _ApplyColorConfMat(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColorConfMat called")
    ' Custom handling for confusion matrix plots if needed
End Sub

Sub _ApplyColor3DScatter(iPlot As Long, RGB_VAL As Long, ALPHA_VAL As Long)
    Const DEBUG_MODE As Boolean = False
    DebugMsg(DEBUG_MODE, "_ApplyColor3DScatter called")
    
    ' Select the plot object
    _SelectGraphObject(iPlot)
    
    ' Apply 3D scatter colors
    With ActiveDocument.CurrentPageItem
        ' Line attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
        
        ' Symbol attributes
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
    End With
End Sub


' ========================================
' Figure Size
' ========================================
Function _cvtMmToInternalUnit(mm As Long)
    Const DEBUG_MODE As Boolean = False
    _cvtMmToInternalUnit = mm*30
End Function

Sub SetFigureSize()
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub

Sub _SetYLabelSize()
    Const DEBUG_MODE As Boolean = False
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub

Sub SetXYLabelSizes()
    Const DEBUG_MODE As Boolean = False
    _SetXLabelSize()
    _SetYLabelSize()
End Sub

' ========================================
' Label Text
' ========================================
Function _SetXLabel()
    Const DEBUG_MODE As Boolean = False
    Dim xLabel As Variant
    xLabel = _ReadCell(GRAPH_PARAMS_COL, X_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, xLabel)
End Function

Function _SetYLabel()
    Const DEBUG_MODE As Boolean = False
    Dim yLabel As Variant
    yLabel = _ReadCell(GRAPH_PARAMS_COL, Y_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, yLabel)
End Function

Sub SetXYLabels()
    Const DEBUG_MODE As Boolean = False
    _SetXLabel()
    _SetYLabel()
End Sub

' ========================================
' Range
' ========================================
Sub _SetXRange()
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
    _SetXRange()
    _SetYRange()
End Sub

' ========================================
' Scales
' ========================================
Function _SetScaleType(axisIndex As Long, scaleType As Long)
    Dim axis As Object
    ' Get the axis object directly
    Set axis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(axisIndex)
    ' Set scale type
    axis.SetAttribute(SAA_TYPE, scaleType)
End Function

Function _cvtScaleTypeFromVariantToLong(cellValue As Variant) As Long
    Const DEBUG_MODE As Boolean = False
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
    _cvtScaleTypeFromVariantToLong = scaleType
End Function

Sub _SetXScale()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim xScaleVariant As Variant
    Dim xScaleLong As Long

    xScaleVariant = _ReadCell(GRAPH_PARAMS_COL, X_SCALE_TYPE_ROW)
    xScaleLong = _cvtScaleTypeFromVariantToLong(xScaleVariant)

    _SetScaleType(HORIZONTAL, xScaleLong)
    On Error GoTo 0
End Sub

Sub _SetYScale()
    Const DEBUG_MODE As Boolean = False
    On Error Resume Next
    Dim yScaleVariant As Variant
    Dim yScaleLong As Long

    yScaleVariant = _ReadCell(GRAPH_PARAMS_COL, Y_SCALE_TYPE_ROW)
    yScaleLong = _cvtScaleTypeFromVariantToLong(yScaleVariant)
    _SetScaleType(VERTICAL, yScaleLong)

    On Error GoTo 0
End Sub

Sub SetScales()
    Const DEBUG_MODE As Boolean = False
    _SetXScale()
    _SetYScale()
End Sub

' ========================================
' Ticks
' ========================================
Sub _SetXTicks()
    Const DEBUG_MODE As Boolean = False
    Dim xTicksFirstRow As Variant
    xTicksFirstRow = _ReadCell(X_TICKS_COL, 0)
    If Not (xTicksFirstRow = "None" Or xTicksFirstRow = "auto") Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, X_TICKS_COL)
    End If
End Sub

Sub _SetYTicks()
    Const DEBUG_MODE As Boolean = False
    Dim yTicksFirstRow As Variant
    yTicksFirstRow = _ReadCell(Y_TICKS_COL, 0)
    If Not (yTicksFirstRow = "None" Or yTicksFirstRow = "auto") Then
        ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOLUSED, 1)
        ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SAA_TICCOL, Y_TICKS_COL)
    End If
End Sub

Sub SetTicks()
    Const DEBUG_MODE As Boolean = False
    _SetXTicks()
    _SetYTicks()
End Sub

' ========================================
' Tick Sizes
' ========================================
Sub _SetXTickSizes()
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
    _SetXTickSizes()
    _SetYTickSizes()
End Sub

' ========================================
' XY and Tick Sizes
' ========================================
Sub SetXYLabelSizesAndTickLabelSizes()
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False
    Dim plotCount As Long
    Dim iPlot As Long
    Dim graphItem As Object
    Dim plotType As String

    ' Get the graph page
    Set graphItem = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
    If graphItem Is Nothing Then
        DebugMsg(DEBUG_MODE, "Error: Graph object not found in SetLineWidth")
        Exit Sub
    End If

    ' Get the number of plots
    plotCount = graphItem.Plots.Count

    ' Loop through all plots and set line width
    For iPlot = 0 To plotCount - 1
        plotType = LCase(_ReadPlotTypeAsStr(iPlot))
        _SelectGraphObject iPlot

        If plotType = "area" Then
           ' ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
           '                                                          SEA_THICKNESS, _
           '                                                          AREA_LINE_THICKNESS)
           ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
                                                                    SEA_THICKNESS, _
                                                                    0)           
        End If
        
        ' For polar plots, set the specific line width
        If plotType = "polar" Then
           ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, _
                                                                    SEA_THICKNESS, _
                                                                    POLAR_LINE_THICKNESS)
        End If
    Next iPlot
End Sub

Sub SetTickLabelRotation()
    Const DEBUG_MODE As Boolean = False
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
    Const DEBUG_MODE As Boolean = False

    ' ' Make sure graph is active
    ' ActiveDocument.NotebookItems(GRAPH_NAME).Open

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

    ' Labels
    SetXYLabels()
    SetXYLabelSizes()

    ' Ticks and Labels
    SetXYLabelSizesAndTickLabelSizes()

    ' Tick label rotation
    SetTickLabelRotation()

    ' Line Width
    SetLineWidth()

    ' Activate the graph page
    ActiveDocument.NotebookItems(GRAPH_NAME).Open

End Sub