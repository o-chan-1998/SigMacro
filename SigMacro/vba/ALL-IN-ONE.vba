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

' ========================================
' Visual Settings
' ========================================
' Display Options
Const HIDE_LEGEND As Long = 0
Const TITLE_HIDE As Long = 0

' Styling
Const TICK_THICKNESS_00 As Variant = &H00000000
Const TICK_THICKNESS_08 As Variant = &H00000008
Const TICK_WIDTH_20 As Variant = &H00000020
Const SSA_COLOR_ALPHA As Long = &H000008a7&

' Font Sizes
Const LABEL_PTS_00 As Variant = "0"
Const LABEL_PTS_07 As String = "97"
Const LABEL_PTS_08 As String = "111"

' ========================================
' Axis Scale Types
' ========================================
Const SAA_TYPE_LINEAR = 1
Const SAA_TYPE_COMMON = 2
Const SAA_TYPE_LOG = 3
Const SAA_TYPE_PROBABILITY = 4
Const SAA_TYPE_PROBIT = 5
Const SAA_TYPE_LOGIT = 6
Const SAA_TYPE_CATEGORY = 7
Const SAA_TYPE_DATETIME = 8

' ========================================
' Worksheet Layout Constants
' ========================================

' For Plot Wizard
' ----------------------------------------
' Columns
Const _PLOT_TYPE_EXPLANATION_COL As Long = 0
Const PLOT_TYPE_COL As Long = 1
' Rows
Const PLOT_TYPE_ROW As Long = 0
Const PLOT_STYLE_ROW As Long = 1
Const PLOT_DATA_TYPE_ROW As Long = 2
Const _PLOT_COLUMNS_PER_PLOT_ROW As Long = 3 ' Spacer
Const _PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW As Long = 4 ' Spacer
Const PLOT_DATA_SOURCE_ROW As Long = 5
Const PLOT_POLARUNITS_ROW As Long = 6
Const PLOT_ANGLEUNITS_ROW As Long = 7
Const PLOT_MIN_ANGLE_ROW As Long = 8
Const PLOT_MAX_ANGLE_ROW As Long = 9
Const PLOT_UNKONWN1_ROW As Long = 10
Const PLOT_GROUP_STYLE_ROW As Long = 11
Const PLOT_USE_AUTOMATIC_LEGENDS_ROW As Long = 12

' Graph Parameters
' ----------------------------------------
' Columns
Const _GRAPH_PARAMS_EXPLANATION_COL As Long = 2
Const GRAPH_PARAMS_COL As Long = 3
' Rows
Const X_LABEL_ROW As Long = 0
Const X_MM_ROW As Long = 1
Const X_SCALE_TYPE_ROW As Long = 2
Const X_MIN_ROW As Long = 3
Const X_MAX_ROW As Long = 4
Const _X_TICKS_ROW As Long = 5
Const Y_LABEL_ROW As Long = 6
Const Y_MM_ROW As Long = 7
Const Y_SCALE_TYPE_ROW As Long = 8
Const Y_MIN_ROW As Long = 9
Const Y_MAX_ROW As Long = 10
Const _Y_TICKS_ROW As Long = 11

' Ticks (Not handled by macros but embedded in JNB file)
' ----------------------------------------
Const _X_TICKS_COL As Long = 4
Const _Y_TICKS_COL As Long = 5

' Data Columns
' ----------------------------------------
' General
Const DATA_START_COL As Long = 16
Const DATA_CHUNK_SIZE As Long = 16
Const DATA_MAX_NUM_PLOTS As Long = 13
' In-chunk indices
Const DATA_ID_X AS LONG = 0
Const DATA_ID_X_ERR AS LONG = 1
Const DATA_ID_X_UPPER AS LONG = 2
Const DATA_ID_X_LOWER AS LONG = 3
Const DATA_ID_Y AS LONG = 4
Const DATA_ID_Y_ERR AS LONG = 5
Const DATA_ID_Y_UPPER AS LONG = 6
Const DATA_ID_Y_LOWER AS LONG = 7
Const DATA_ID_RGBA AS LONG = 8

' ========================================
' Helper
' ========================================
Sub DebugMsg(msg As String)
    MsgBox msg, vbInformation, "Debug Info"
End Sub

Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object
    Dim cellValue As Variant
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function


Sub SelectGraphObject(plotIndex As Long)
    On Error Resume Next

    Dim plotObj As Object
    Set plotObj = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not plotObj Is Nothing Then
        plotObj.SetObjectCurrent
        If Err.Number <> 0 Then
            ' DebugMsg "Error in SelectGraphObject: " & Err.Description
            Err.Clear
        End If
    Else
        DebugMsg "Plot object not found in SelectGraphObject for index " & plotIndex
    End If
End Sub


Sub Plot()
    ' Open the worksheet
    ActiveDocument.NotebookItems(WORKSHEET_NAME).Open
    ActiveDocument.CurrentDataItem.Open
    
    ' Get plot configuration
    Dim PLOT_TYPE As String
    Dim plotStyle As String
    Dim dataType As String
    Dim columnsPerPlot As Variant
    Dim plotColumnsCountArray As Variant
    Dim dataSource As String
    Dim polarUnits As String
    Dim angleUnits As String
    Dim minAngle As Double
    Dim maxAngle As Double
    Dim unknown1 As Variant
    Dim groupStyle As Variant
    Dim useAutomaticLegends As Variant
    
    PLOT_TYPE = _ReadCell(PLOT_TYPE_COL, PLOT_TYPE_ROW)
    plotStyle = _ReadCell(PLOT_TYPE_COL, PLOT_STYLE_ROW)
    dataType = _ReadCell(PLOT_TYPE_COL, PLOT_DATA_TYPE_ROW)
    columnsPerPlot = _ReadCell(PLOT_TYPE_COL, PLOT_COLUMNS_PER_PLOT_ROW)
    plotColumnsCountArray = _ReadCell(PLOT_TYPE_COL, PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW)
    dataSource = _ReadCell(PLOT_TYPE_COL, PLOT_DATA_SOURCE_ROW)
    polarUnits = _ReadCell(PLOT_TYPE_COL, PLOT_POLARUNITS_ROW)
    angleUnits = _ReadCell(PLOT_TYPE_COL, PLOT_ANGLEUNITS_ROW)
    minAngle = _ReadCell(PLOT_TYPE_COL, PLOT_MIN_ANGLE_ROW)
    maxAngle = _ReadCell(PLOT_TYPE_COL, PLOT_MAX_ANGLE_ROW)
    unknown1 = _ReadCell(PLOT_TYPE_COL, PLOT_UNKONWN1_ROW)
    groupStyle = _ReadCell(PLOT_TYPE_COL, PLOT_GROUP_STYLE_ROW)
    useAutomaticLegends = _ReadCell(PLOT_TYPE_COL, PLOT_USE_AUTOMATIC_LEGENDS_ROW)
    
    ' Build column arrays dynamically based on constants
    Dim i As Long
    Dim currentColumn As Long
    Dim graphAlreadyExists As Boolean
    graphAlreadyExists = _GetGraphObject()
    currentColumn = DATA_START_COL
    
    For i = 0 To DATA_MAX_NUM_PLOTS - 1
        ' Get column mapping and count for current plot type
        Dim ColumnsPerPlot As Variant
        Dim PlotColumnCountArray As Variant
        ColumnsPerPlot = _GetColumnMapping(PLOT_TYPE, currentColumn)
        PlotColumnCountArray = _GetColumnCount(PLOT_TYPE)
        
        ' Increment currentColumn
        currentColumn = currentColumn + DATA_CHUNK_SIZE
        
        ' Create the plot if not exists
        If Not graphAlreadyExists And i = 0 Then
            ' First plot with no existing graph - create the graph
            ActiveDocument.CurrentPageItem.CreateWizardGraph(PLOT_TYPE, _
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
                unknown1, _
                groupStyle, _
                useAutomaticLegends)
            graphAlreadyExists = True
        Else
            ' If graph exists, add plot
            ActiveDocument.NotebookItems(GRAPH_NAME).Open
            ActiveDocument.CurrentPageItem.AddWizardPlot(PLOT_TYPE, _
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
                , _
                groupStyle, _
                useAutomaticLegends)
        End If
    Next i
    
    ' Open the graph and select the plot
    ActiveDocument.NotebookItems(GRAPH_NAME).Open
End Sub

' ========================================
' Removers
' ========================================
Sub RemoveLegend()
   ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
   ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_AUTOLEGENDSHOW, HIDE_LEGEND)
End Sub

Sub RemoveXSpines()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_00)
End Sub

Sub RemoveYSpines()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_00)
End Sub

Sub RemoveTitle()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_FLAG_TITLESUNALIGNED, TITLE_HIDE)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_TITLESHOW, TITLE_HIDE)
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
            _DetectPlotType = "Line/Scatter"
        Case SLA_TYPE_BAR
            _DetectPlotType = "Bar"
        Case SLA_TYPE_STACKED
            _DetectPlotType = "Stacked Bar"
        Case SLA_TYPE_TUKEY
            _DetectPlotType = "Box"
        Case SLA_TYPE_3DSCATTER
            _DetectPlotType = "3D Scatter/Line"
        Case SLA_TYPE_MESH
            _DetectPlotType = "MESH"
        Case SLA_TYPE_CONTOUR
            _DetectPlotType = "CONTOUR"
        Case SLA_TYPE_POLAR
            _DetectPlotType = "POLAR"
        Case SLA_TYPE_MAXVAL
            _DetectPlotType = "MAXVAL"
        Case Else
            _DetectPlotType = "Unknown Object Type: " & object_type
    End Select
    
    ' DebugMsg "Type Detected: " & _DetectPlotType
    Exit Function
ErrorHandler:
    _DetectPlotType = "An error has occurred: " & Err.Description
End Function

Sub _ChangeColorLine(RGB_VAL As Long, plotIndex As Long)
    ' SEA = Set Line Attribute
    ' DebugMsg "_ChangeColorLine called"
    SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, RGB_VAL)
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORCOL, -2)       
       ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, 2)       
       .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SOA_COLOR, RGB_VAL)
    End With
End Sub

Sub _ChangeColorSymbol(RGB_VAL As Long, ALPHA_VAL As Long, plotIndex As Long)
    ' SSA = Set Symbol Attribute
    ' DebugMsg "_ChangeColorSymbol called"
    SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLORREPEAT, 2)        
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR_ALPHA, ALPHA_VAL)
    End With
End Sub

Sub _ChangeColorSolid(RGB_VAL As Long, plotIndex As Long)
    ' SDA = Set Solid Attribute
    ' Solids include graph planes, bars, and drawn solids objects
    ' DebugMsg "_ChangeColorSolid called"
    SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)        
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLORREPEAT, 2)
    End With
End Sub

Sub _ChangeColorErrorBar(RGB_VAL As Long, plotIndex As Long)
    ' SLA = Set Line Attributes
    SelectGraphObject plotIndex
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
    SelectGraphObject plotIndex
    With ActiveDocument.CurrentPageItem
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, RGB_VAL)
        ' .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, 2)
        .SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_EDGECOLOR, RGB_BLACK)
    End With
End Sub

Function _CalculateColorColumnForPlot(plotIndex As Long) As Long
    _CalculateColorColumnForPlot = DATA_START_COL + (plotIndex * DATA_CHUNK_SIZE) + DATA_ID_RGBA
End Function

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
    Dim i As Long
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
    For i = 0 To plotCount - 1
        colorColumn = _CalculateColorColumnForPlot(i)
        RGB_VAL = _GetRGBFromColumn(colorColumn)
        ALPHA_VAL = _GetAlphaFromColumn(colorColumn)

        ' Apply color based on plot type
        Select Case DetectedPlotType
            Case "Line/Scatter"
                _ChangeColorLine RGB_VAL, i
                _ChangeColorSymbol RGB_VAL, ALPHA_VAL, i
                _ChangeColorSolid RGB_VAL, i
            Case "3DScatter"
                _ChangeColorLine RGB_VAL, i
                _ChangeColorSymbol RGB_VAL, ALPHA_VAL, i
            Case "Bar", "Stacked"
                _ChangeColorSolid RGB_VAL, i
            Case "Box"
                _ChangeColorBox RGB_VAL, i
            Case Else
                DebugMsg "Unknown plot type detected: " & DetectedPlotType
        End Select
    Next i
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
   ' Width   
   Dim xLength_mm As Long
   Dim xLength_sp As Long
   xLength_mm = _ReadCell(GRAPH_PARAMS_COL, X_MM_ROW)
   xLength_sp = _cvtMmToInternalUnit(xLength_mm)

   ' Height
   Dim yLength_mm As Long
   Dim yLength_sp As Long
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

Sub SetLabelSizes()
   On Error Resume Next
   _SetXLabelSize()
   _SetYLabelSize()
   On Error GoTo 0
End Sub

' ========================================
' Label Text
' ========================================
Function _SetXLabel()
	Dim xLabel As Variant
	xLabel = _ReadCell(GRAPH_PARAMS_COL, X_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X) ' Select X-axis
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, xLabel) ' Set X-axis label
End Function

Function _SetYLabel()
	Dim yLabel As Variant
	yLabel = _ReadCell(GRAPH_PARAMS_COL, Y_LABEL_ROW)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y) ' Select Y-axis
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, yLabel) ' Set Y-axis label
End Function

Sub SetLabels()
    _SetXLabel() 
    _SetYLabel() 
End Sub

' ========================================
' Range
' ========================================
Sub _SetXRange()
    Dim xMin As Variant, xMax As Variant
    Dim xAxis As Object
    
    xMin = _ReadCell(GRAPH_PARAMS_COL, X_MIN_ROW)
    xMax = _ReadCell(GRAPH_PARAMS_COL, X_MAX_ROW)
    
    ' Get the X axis object directly
    Set xAxis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(AXIS_X)
    
    ' Set the min and max values
    xAxis.SetAttribute(SAA_FROMVAL, xMin)
    xAxis.SetAttribute(SAA_TOVAL, xMax)
End Sub

Sub _SetYRange()
    Dim yMin As Variant, yMax As Variant
    Dim yAxis As Object
    
    yMin = _ReadCell(GRAPH_PARAMS_COL, Y_MIN_ROW)
    yMax = _ReadCell(GRAPH_PARAMS_COL, Y_MAX_ROW)

    ' Get the Y axis object directly
    Set yAxis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(AXIS_Y)

    ' Set the min and max values
    yAxis.SetAttribute(SAA_FROMVAL, yMax)
    yAxis.SetAttribute(SAA_TOVAL, yMin)
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

Sub SetScales()
   On Error Resume Next
   Dim xScaleData As Variant, yScaleData As Variant
   Dim xScaleType As Long, yScaleType As Long

   ' Read scale types from worksheet
   xScaleData = _ReadCell(GRAPH_PARAMS_COL, X_SCALE_TYPE_ROW)
   yScaleData = _ReadCell(GRAPH_PARAMS_COL, Y_SCALE_TYPE_ROW)   

   ' Convert to scale type constants
   xScaleType = _GetScaleType(xScaleData)
   yScaleType = _GetScaleType(yScaleData)

   ' Set X axis scale type
   _SetAxisType AXIS_X, xScaleType
   If Err.Number <> 0 Then
      DebugMsg "Error setting X axis scale: " & Err.Description
      Err.Clear
   End If

   ' Set Y axis scale type
   _SetAxisType AXIS_Y, yScaleType
   If Err.Number <> 0 Then
      DebugMsg "Error setting Y axis scale: " & Err.Description
      Err.Clear
   End If

On Error GoTo 0
End Sub

' ========================================
' Tick Sizes
' ========================================
Sub _SetXTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_THICKNESS_08)
End Sub

Sub _SetYTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_THICKNESS_08)
End Sub

Sub SetTickSizes()
   On Error Resume Next
   _SetXTickSizes()
   _SetYTickSizes()
   On Error GoTo 0
End Sub

' ========================================
' Plot
' ========================================
Function _GetColumnMapping(plotType As String, currentColumn As Long) As Variant
    Dim mapping()
    Select Case plotType
        Case "Vertical Bar Chart"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + DATA_ID_X        ' X column
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Y column
            mapping(0, 2) = currentColumn + DATA_ID_Y_ERR    ' Y error column
        Case "Horizontal Bar Chart"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + DATA_ID_Y        ' Y column (category)
            mapping(0, 1) = currentColumn + DATA_ID_X        ' X column (value)
            mapping(0, 2) = currentColumn + DATA_ID_X_ERR    ' X error column
        Case "Line Plot", "Scatter Plot"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + DATA_ID_X        ' X column
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Y column
            mapping(0, 2) = currentColumn + DATA_ID_Y_ERR    ' Y error column
        Case "Area Plot"
            ReDim mapping(2, 3)
            mapping(0, 0) = currentColumn + DATA_ID_X        ' X column
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Y column
            mapping(0, 2) = currentColumn + DATA_ID_Y_LOWER  ' Y lower bound
            mapping(0, 3) = currentColumn + DATA_ID_Y_UPPER  ' Y upper bound
        Case "Box Plot"
            ReDim mapping(2, 1)
            mapping(0, 0) = currentColumn + DATA_ID_X        ' X column (category)
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Y column (values)
        Case "Horizontal Box Plot"                            
            ReDim mapping(2, 1)                               
            mapping(0, 0) = currentColumn + DATA_ID_Y        ' Y column (category)
            mapping(0, 1) = currentColumn + DATA_ID_X        ' X column (values)
        Case "Violin Plot"                                    
            ReDim mapping(2, 1)                               
            mapping(0, 0) = currentColumn + DATA_ID_X        ' X column (category)
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Y column (values)
        Case "Polar Plot"                                     
            ReDim mapping(2, 1)                               
            mapping(0, 0) = currentColumn + DATA_ID_X        ' Theta column
            mapping(0, 1) = currentColumn + DATA_ID_Y        ' Radius column
        Case Else  ' Default to Vertical Bar Chart
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + DATA_ID_X
            mapping(0, 1) = currentColumn + DATA_ID_Y
            mapping(0, 2) = currentColumn + DATA_ID_Y_ERR
    End Select
    ' Fill in the row ranges for all columns
    Dim i As Integer, j As Integer
    For i = 0 To UBound(mapping, 2)
        mapping(1, i) = 0          ' Start row
        mapping(2, i) = 31999999   ' End row
    Next i
    _GetColumnMapping = mapping
End Function

' Get column count based on plot type
Function _GetColumnCount(plotType As String) As Variant
    Dim countArray()
    ReDim countArray(0)
    Select Case plotType
        Case "Vertical Bar Chart", "Horizontal Bar Chart", "Line Plot", "Scatter Plot"
            countArray(0) = 3
        Case "Area Plot"
            countArray(0) = 4
        Case "Box Plot", "Horizontal Box Plot", "Violin Plot", "Polar Plot"
            countArray(0) = 2
        Case Else
            countArray(0) = 3
    End Select
    _GetColumnCount = countArray
End Function

Function _GetGraphObject() As Boolean
    On Error Resume Next
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Dim tempGraph As Object
        Set tempGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not tempGraph Is Nothing Then
            _GetGraphObject = True
            Exit Function
        End If
    End If
    _GetGraphObject = False
    DebugMsg "Error: Could not locate or create the graph object"
End Function




' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   Plot()   
   RemoveLegend()
   RemoveXSpines()
   RemoveYSpines()
   RemoveTitle()
   SetColors()
   SetFigureSize()
   SetLabelSizes()
   SetLabels()
   SetRanges()
   SetScales()
   SetTickSizes()
   On Error GoTo 0
End Sub