Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const WORKSHEET_NAME As String = "worksheet"
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const X_LABEL_COL = 0
Const Y_LABEL_COL = 6

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Function SetXLabel()
	Dim xLabel As Variant
	xLabel = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(X_LABEL_COL, 0, X_LABEL_COL, 0)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X) ' Select X-axis
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, xLabel(0,0)) ' Set X-axis label
End Function

Function SetYLabel()
	Dim yLabel As Variant
	yLabel = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(Y_LABEL_COL, 0, Y_LABEL_COL, 0)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y) ' Select Y-axis
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_RTFNAME, yLabel(0,0)) ' Set Y-axis label
End Function

' ========================================
' Main
' ========================================
Sub Main()
    SetXLabel
    SetYLabel
End Sub