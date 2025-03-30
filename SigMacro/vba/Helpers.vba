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
