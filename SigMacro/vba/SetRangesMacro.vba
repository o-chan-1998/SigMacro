Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const WORKSHEET_NAME As String = "worksheet"
Const AXIS_X As Long = 0
Const AXIS_Y As Long = 1
Const X_MIN_COL = 3
Const X_MAX_COL = 4
Const Y_MIN_COL = 10
Const Y_MAX_COL = 11

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Function SetXRange()
    Dim xMin As Variant, xMax As Variant
    Dim xAxis As Object
    
    xMin = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(X_MIN_COL, 0, X_MIN_COL, 0)
    xMax = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(X_MAX_COL, 0, X_MAX_COL, 0)
    
    ' Get the X axis object directly
    Set xAxis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(AXIS_X)
    
    ' Set the min and max values
    xAxis.SetAttribute(SAA_FROMVAL, xMin(0,0))
    xAxis.SetAttribute(SAA_TOVAL, xMax(0,0))
End Function

Function SetYRange()
    Dim yMin As Variant, yMax As Variant
    Dim yAxis As Object
    
    yMin = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(Y_MIN_COL, 0, Y_MIN_COL, 0)
    yMax = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(Y_MAX_COL, 0, Y_MAX_COL, 0)
    
    ' Get the Y axis object directly
    Set yAxis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(AXIS_Y)

    ' Set the min and max values
    yAxis.SetAttribute(SAA_FROMVAL, yMax(0,0))
    yAxis.SetAttribute(SAA_TOVAL, yMin(0,0))
    
End Function

' ========================================
' Main
' ========================================
Sub Main()
    SetXRange
    SetYRange
End Sub
