Option Explicit
' ----------------------------------------
' Note
' ----------------------------------------
' For compatibility, graphs are arranged in chunks of "COLUMN_SPACING" columns:
'     x1, xerr1, x_lower1, x_upper1, y1, yerr1, y_lower1, y_upper1, rgba1,
'     x2, xerr2, x_lower2, x_upper2, y2, yerr2, y_lower2, y_upper2, rgba2,
'     ...
'     xN, xerrN, x_lowerN, x_upperN, yN, yerrN, y_lowerN, y_upperN, rgbaN
'
' However, only some columns are used. For example, with "Vertical Bar Chart" type,
' only xK (column 0 in the chunk), yK (column 2 in the chunk), and yerrK (column 3 in the chunk) are used.

' ----------------------------------------
' Constants
' ----------------------------------------
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"
Const FIRST_DATA_COLUMN As Long = 32
Const COLUMN_SPACING As Long = 9
Const NUM_PLOTS As Long = 13

' For Plot Wizard
Const PLOT_TYPE_COL As Long = 0
Const PLOT_TYPE_ROW As Long = 0
Const PLOT_STYLE_ROW As Long = 1
Const PLOT_DATA_TYPE_ROW As Long = 2
Const _PLOT_COLUMNS_PER_PLOT_ROW As Long = 3 ' Spacer
Const _PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW As Long = 4 ' Spacer
Const PLOT_DATA_SOURCE_ROW As Long = 5
Const PLOT_POLARUNITS_ROW As Long = 6
Const PLOT_ANGLEUNITS_ROW As Long = 7
Const PLOT_MIN_ANGLE_ROW As Long = 8
Const PLOT_MAX_ANGLE_ROW As Long = 8
Const PLOT_UNKONWN1_ROW As Long = 9
Const PLOT_GROUP_STYLE_ROW As Long = 10
Const PLOT_USE_AUTOMATIC_LEGENDS_ROW As Long = 11

' For data
Const XIdxInChunk As Long = 0
Const XErrIdxInChunk As Long = 1
Const XUpperIdxInChunk As Long = 2
Const XLowerIdxInChunk As Long = 3
Const YIdxInChunk As Long = 4
Const YErrIdxInChunk As Long = 5
Const YUpperIdxInChunk As Long = 6
Const YLowerIdxInChunk As Long = 7
Const RGBAIdxInChunk As Long = 8


' ----------------------------------------
' Functions
' ----------------------------------------
Sub DebugMsg(msg As String)
    MsgBox msg, vbInformation, "Debug Info"
End Sub


Function ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim cellValue As Variant
    cellValue = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    ReadCell = cellValue(0, 0)
End Function

' ----------------------------------------
' Plot Configuration
' ----------------------------------------

' Get column mapping based on plot type
Function GetColumnMapping(plotType As String, currentColumn As Long) As Variant
    Dim mapping()
    Select Case plotType
        Case "Vertical Bar Chart"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' X column
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Y column
            mapping(0, 2) = currentColumn + YErrIdxInChunk    ' Y error column
        Case "Horizontal Bar Chart"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + YIdxInChunk       ' Y column (category)
            mapping(0, 1) = currentColumn + XIdxInChunk       ' X column (value)
            mapping(0, 2) = currentColumn + XErrIdxInChunk    ' X error column
        Case "Line Plot", "Scatter Plot"
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' X column
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Y column
            mapping(0, 2) = currentColumn + YErrIdxInChunk    ' Y error column
        Case "Area Plot"
            ReDim mapping(2, 3)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' X column
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Y column
            mapping(0, 2) = currentColumn + YLowerIdxInChunk  ' Y lower bound
            mapping(0, 3) = currentColumn + YUpperIdxInChunk  ' Y upper bound
        Case "Box Plot"
            ReDim mapping(2, 1)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' X column (category)
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Y column (values)
        Case "Horizontal Box Plot"
            ReDim mapping(2, 1)
            mapping(0, 0) = currentColumn + YIdxInChunk       ' Y column (category)
            mapping(0, 1) = currentColumn + XIdxInChunk       ' X column (values)
        Case "Violin Plot"
            ReDim mapping(2, 1)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' X column (category)
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Y column (values)
        Case "Polar Plot"
            ReDim mapping(2, 1)
            mapping(0, 0) = currentColumn + XIdxInChunk       ' Theta column
            mapping(0, 1) = currentColumn + YIdxInChunk       ' Radius column
        Case Else  ' Default to Vertical Bar Chart
            ReDim mapping(2, 2)
            mapping(0, 0) = currentColumn + XIdxInChunk
            mapping(0, 1) = currentColumn + YIdxInChunk
            mapping(0, 2) = currentColumn + YErrIdxInChunk
    End Select
    ' Fill in the row ranges for all columns
    Dim i As Integer, j As Integer
    For i = 0 To UBound(mapping, 2)
        mapping(1, i) = 0          ' Start row
        mapping(2, i) = 31999999   ' End row
    Next i
    GetColumnMapping = mapping
End Function

' Get column count based on plot type
Function GetColumnCount(plotType As String) As Variant
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
    GetColumnCount = countArray
End Function

Function GetGraphObject() As Boolean
    On Error Resume Next
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Dim tempGraph As Object
        Set tempGraph = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not tempGraph Is Nothing Then
            GetGraphObject = True
            Exit Function
        End If
    End If
    GetGraphObject = False
    DebugMsg "Error: Could not locate or create the graph object"
End Function

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
    Dim additionalParam1 As Variant
    Dim additionalParam2 As Variant
    Dim additionalParam3 As Variant
    
    PLOT_TYPE = ReadCell(PLOT_TYPE_COL, PLOT_TYPE_ROW)
    plotStyle = ReadCell(PLOT_TYPE_COL, PLOT_STYLE_ROW)
    dataType = ReadCell(PLOT_TYPE_COL, PLOT_DATA_TYPE_ROW)
    columnsPerPlot = ReadCell(PLOT_TYPE_COL, PLOT_COLUMNS_PER_PLOT_ROW)
    plotColumnsCountArray = ReadCell(PLOT_TYPE_COL, PLOT_PLOT_COLUMNS_COUNT_ARRAY_ROW)
    dataSource = ReadCell(PLOT_TYPE_COL, PLOT_DATA_SOURCE_ROW)
    polarUnits = ReadCell(PLOT_TYPE_COL, PLOT_POLARUNITS_ROW)
    angleUnits = ReadCell(PLOT_TYPE_COL, PLOT_ANGLEUNITS_ROW)
    minAngle = ReadCell(PLOT_TYPE_COL, PLOT_MIN_ANGLE_ROW)
    maxAngle = ReadCell(PLOT_TYPE_COL, PLOT_MAX_ANGLE_ROW)
    unkonwn1 = ReadCell(PLOT_TYPE_COL, PLOT_UNKONWN1_ROW)
    groupStyle = ReadCell(PLOT_TYPE_COL, PLOT_GROUP_STYLE_ROW)
    UseAutomaticLegends = ReadCell(PLOT_TYPE_COL, PLOT_USE_AUTOMATIC_LEGENDS_ROW)
    
    ' Build column arrays dynamically based on constants
    Dim i As Long
    Dim currentColumn As Long
    Dim graphAlreadyExists As Boolean
    graphAlreadyExists = GetGraphObject()
    currentColumn = FIRST_DATA_COLUMN
    
    For i = 0 To NUM_PLOTS - 1
        ' Get column mapping and count for current plot type
        Dim ColumnsPerPlot As Variant
        Dim PlotColumnCountArray As Variant
        ColumnsPerPlot = GetColumnMapping(PLOT_TYPE, currentColumn)
        PlotColumnCountArray = GetColumnCount(PLOT_TYPE)
        
        ' Increment currentColumn
        currentColumn = currentColumn + COLUMN_SPACING
        
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

Sub Main()
   Plot()
End Sub