Option Explicit
' ========================================
' General Constants
' ========================================
Const DEBUG_MODE As Boolean = False
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"

Function FindColumn(oWorksheet As Object, columnName As String) As Long
    Dim ColIndex As Long
    Dim i As Long
    
    ColIndex = -1 ' Default return value if column not found
    
    ' Loop through all columns in current worksheet
    For i = 1 To oWorksheet.Columns.Count
        ' Compare column name (case-insensitive)
        If LCase(oWorkSheet.Columns(i).Label) = LCase(ColumnName) Then
            ColIndex = i
            Exit For
        End If
    Next i
    
    FindColumn = ColIndex
End Function


Sub Main()
    Dim oWorksheet As Object
    Set oWorksheet = ActiveDocument.NotebookItems(WORKSHEET_NAME)

    Dim oDataTable As Object
    Set oDataTable = oWorkSHeet.DataTable

    Dim foundCol As Long    
    foundCol = FindColumn(oWorksheet, "xticks")
   
    MsgBox "Development: " & foundCol
End Sub
