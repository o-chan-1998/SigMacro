Option Explicit


' ----------------------------------------
' Constants
' ----------------------------------------
Const DEBUG_MODE As Boolean = False
Const WORKSHEET_NAME As String = "worksheet"
Const COLOR_COLUMN_FIRST As Long = 16
Const COLOR_COLUMN_SPACING As Long = 5
Const SSA_COLOR_ALPHA As Long = &H000008a7&

' ----------------------------------------
' Functions
' ----------------------------------------
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
        MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Function GetRGBFromColumn(columnIndex As Long) As Long
    ' DebugMsg "GetRGBFromColumn called for plot " & columnIndex
   Dim rValue As Variant, gValue As Variant, bValue As Variant
   
    ' Read RGB values from worksheet (R, G, B values are assumed to be in adjacent columns)
    rValue = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(columnIndex, 2, columnIndex, 2)    
    gValue = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(columnIndex, 1, columnIndex, 1)
    bValue = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(columnIndex, 0, columnIndex, 0)    

    ' Convert to integers and create RGB color
    Dim r As Integer, g As Integer, b As Integer
    r = CInt(rValue(0, 0))
    g = CInt(gValue(0, 0))
    b = CInt(bValue(0, 0))

    ' Standard RGB (VBA default)
    GetRGBFromColumn = RGB(r, g, b)
End Function

Function GetAlphaFromColumn(columnIndex As Long) As Long
    ' DebugMsg "GetRGBFromColumn called for plot " & columnIndex
   Dim alphaValue As Variant
   
    ' Read RGB values from worksheet (R, G, B values are assumed to be in adjacent columns)
    alphaValue = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(columnIndex, 3, columnIndex, 3)

    ' Standard RGB (VBA default)
    GetAlphaFromColumn = alphaValue
End Function


Function CalculateColorColumnForPlot(plotIndex As Long) As Long
    CalculateColorColumnForPlot = COLOR_COLUMN_FIRST + (plotIndex * COLOR_COLUMN_SPACING)
End Function

Function SelectPlotObject(plotIndex As Long) As Object
    On Error Resume Next

    Set SelectPlotObject = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(plotIndex)
    If Not SelectPlotObject Is Nothing Then
        On Error Resume Next
        SelectPlotObject.SetObjectCurrent
        If Err.Number <> 0 Then
            DebugMsg "Error setting plot " & plotIndex & " as current: " & Err.Description
            Err.Clear
        End If
    Else
        Dim plotCount As Long
        plotCount = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots.Count
    End If
End Function

Function DetectPlotType() As String
    On Error GoTo ErrorHandler
    Dim ObjectType As Variant
    Dim object_type As Variant
    object_type = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(0).GetAttribute(SLA_TYPE, ObjectType)
    If object_type = False Then
        DetectPlotType = "Error: Failed to get object type."
        Exit Function
    End If

    ' Mapping
    Select Case object_type
        Case SLA_TYPE_SCATTER, SLA_TYPE_MINVAL, SLA_TYPE_POLARXY, SLA_TYPE_3DBAR, SLA_TYPE_TERNARYSCATTER
            DetectPlotType = "Line/Scatter"
        Case SLA_TYPE_BAR
            DetectPlotType = "Bar"
        Case SLA_TYPE_STACKED
            DetectPlotType = "Stacked Bar"
        Case SLA_TYPE_TUKEY
            DetectPlotType = "Box"
        Case SLA_TYPE_3DSCATTER
            DetectPlotType = "3D Scatter/Line"
        Case SLA_TYPE_MESH
            DetectPlotType = "MESH"
        Case SLA_TYPE_CONTOUR
            DetectPlotType = "CONTOUR"
        Case SLA_TYPE_POLAR
            DetectPlotType = "POLAR"
        Case SLA_TYPE_MAXVAL
            DetectPlotType = "MAXVAL"
        Case Else
            DetectPlotType = "Unknown Object Type: " & object_type
    End Select
    
    ' DebugMsg "Type Detected: " & DetectPlotType
    Exit Function
ErrorHandler:
    DetectPlotType = "An error has occurred: " & Err.Description
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
' ----------------------------------------
' Color Setters
' ----------------------------------------
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

' ----------------------------------------
' Main
' ----------------------------------------
Sub Main()
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
    DetectedPlotType = DetectPlotType()
    
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
        colorColumn = CalculateColorColumnForPlot(i)
        RGB_VAL = GetRGBFromColumn(colorColumn)
        ALPHA_VAL = GetAlphaFromColumn(colorColumn)

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