Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const WORKSHEET_NAME As String = "worksheet"
Const X_MM_COL = 1
Const Y_MM_COL = 7

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Function cvtMmToInternalUnit(mm As Long)
   cvtMmToInternalUnit = mm*30
End Function

Sub setFigureSize()
   ' Width   
   Dim xLengthCell As Variant
   Dim xLength_mm As Long
   Dim xLength_sp As Long
   xLengthCell = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(X_MM_COL, 0, X_MM_COL, 0)
   xLength_mm = xLengthCell(0,0)
   xLength_sp = cvtMmToInternalUnit(xLength_mm)

   ' Height
   Dim yLengthCell As Variant
   Dim yLength_mm As Long
   Dim yLength_sp As Long
   yLengthCell = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(Y_MM_COL, 0, Y_MM_COL, 0)
   yLength_mm = yLengthCell(0,0)
   yLength_sp = cvtMmToInternalUnit(yLength_mm)

   With ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH)
      .Width = xLength_sp
      .Height = yLength_sp
   End With
End Sub

' ========================================
' Main
' ========================================
Sub Main()
    setFigureSize
End Sub