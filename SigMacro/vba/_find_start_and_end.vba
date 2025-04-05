Option Explicit
' ========================================
' General Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const WORKSHEET_NAME As String = "worksheet"
Const GRAPH_NAME As String = "graph"
Const ChunkStartColumnNameBase As String = "gw_param_keys "
Const LABEL_ROW As Long = -1

Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object, cellValue As Variant
    Set dataTable = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function

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
    colName = ChunkStartColumnNameBase & iPlot
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

Sub Main()
    Dim startCol As Long, endCol As Long, ii As Long
    For ii = 0 To 10
       startCol = _FindChunkStartCol(ii)
       endCol = _FindChunkEndCol(ii)       
       ' This is the start column of the iPlot = ii
       MsgBox "Development: " & startCol & "-" & endCol
    Next ii
End Sub
