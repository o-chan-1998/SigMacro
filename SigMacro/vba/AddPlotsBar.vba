Option Explicit
' ----------------------------------------
' Note
' ----------------------------------------
' For compatibility, graphs are arranged in chunks of 5 columns:
' x1, xerr1, y1, yerr1, rgba1, x2, xerr2, y2, yerr2, rgba2, ..., xN, xerrN, yN, yerrN, rgbaN
' However, only some columns are used. For example, with "Vertical Bar Chart" type,
' only xK (column 0 in the chunk), yK (column 2 in the chunk), and yerrK (column 3 in the chunk) are used.

' ----------------------------------------
' Constants
' ----------------------------------------
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"
' Constants for plot configuration
Const PLOT_TYPE As String = "Vertical Bar Chart"
Const PLOT_STYLE As String = "Simple Error Bars"
Const DATA_TYPE As String = "XY Pair"
Const DATA_SOURCE As String = "Worksheet Columns"
Const POLAR_UNITS As String = "None"
Const ANGLE_UNITS As String = "Degrees"
Const MIN_ANGLE As Double = 0.0
Const MAX_ANGLE As Double = 360.0
Const GROUP_STYLE As String = "None"
Const USE_AUTOMATIC_LEGENDS As Boolean = True
' Constants for data columns
Const FIRST_DATA_COLUMN As Long = 12
Const COLUMNS_PER_PLOT As Long = 3
Const NUM_PLOTS As Long = 1
Const COLUMN_SPACING As Long = 5
' ----------------------------------------
' Functions
' ----------------------------------------
Sub DebugMsg(msg As String)
    MsgBox msg, vbInformation, "Debug Info"
End Sub

Function GetGraphObject() As Object
    On Error Resume Next
    ' Try several methods to get the graph object
    ' Method 1: Using graph name and notebook items
    Dim graphObj As Object
    Set graphObj = ActiveDocument.NotebookItems(GRAPH_NAME)
    If Not graphObj Is Nothing Then
        graphObj.Open
        Set GetGraphObject = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
        If Not GetGraphObject Is Nothing Then Exit Function
    End If
    ' If we reach here, we couldn't find the graph
    DebugMsg "Error: Could not locate the graph object"
End Function

Sub Main()
    ' Open the worksheet
    ' ActiveDocument.NotebookItems(WORKSHEET_NAME).CurrentDataItem
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    
    ' Build column arrays dynamically based on constants
    Dim i As Long
    Dim currentColumn As Long


    ' Set column data with spacing of 5 columns
    currentColumn = FIRST_DATA_COLUMN
    For i = 0 To 13
		Dim ColumnsPerPlot()
		ReDim ColumnsPerPlot(2, 2)

    	ColumnsPerPlot(0, 0) = currentColumn
		ColumnsPerPlot(1, 0) = 0
		ColumnsPerPlot(2, 0) = 31999999
		ColumnsPerPlot(0, 1) = currentColumn + 2
		ColumnsPerPlot(1, 1) = 0
		ColumnsPerPlot(2, 1) = 31999999
		ColumnsPerPlot(0, 2) = currentColumn + 3
		ColumnsPerPlot(1, 2) = 0
		ColumnsPerPlot(2, 2) = 31999999
        currentColumn = currentColumn + COLUMN_SPACING

        Dim PlotColumnCountArray()
		ReDim PlotColumnCountArray(0)
		PlotColumnCountArray(0) = 3

    	' Create the plot
    	ActiveDocument.CurrentPageItem.AddWizardPlot(PLOT_TYPE, _
												 PLOT_STYLE, _
												DATA_TYPE, _
                                                ColumnsPerPlot, _
												PlotColumnCountArray, _
                                                DATA_SOURCE, _
												POLAR_UNITS, _
												ANGLE_UNITS, _
                                                MIN_ANGLE, _
												MAX_ANGLE, _
												, _
												GROUP_STYLE, _
                                                USE_AUTOMATIC_LEGENDS)

    Next i
    

    ' Open the graph and select the plot
    ActiveDocument.NotebookItems(GRAPH_NAME).Open

End Sub