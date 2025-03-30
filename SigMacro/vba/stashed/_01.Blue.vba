Option Explicit

Function getColor(colorName As String) As Long
    Select Case colorName

        Case "Black"
            getColor = RGB(0, 0, 0)
        Case "Gray"
            getColor = RGB(128, 128, 128)            
        Case "White"
            getColor = RGB(255, 255, 255)    
    
    
        Case "Blue"
            getColor = RGB(0, 128, 192)
        Case "Green"
            getColor = RGB(20, 180, 20)            
        Case "Red"
            getColor = RGB(255, 70, 50)
    
    
        Case "Yellow"
            getColor = RGB(230, 160, 20)
        Case "Purple"
            getColor = RGB(200, 50, 255)
    
    
        Case "Pink"
            getColor = RGB(255, 150, 200)            
        Case "LightBlue"
            getColor = RGB(20, 200, 200)
    
    
        Case "DarkBlue"
            getColor = RGB(0, 0, 100)
        Case "Dan"
            getColor = RGB(228, 94, 50)
        Case "Brown"
            getColor = RGB(128, 0, 0)            

        Case Else
            ' Default or error handling
            getColor = RGB(0, 0, 0)
    End Select
End Function


Sub updatePlot(COLOR As Long)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, COLOR)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_THICKNESS, &H00000005&)
    ' MsgBox "updatePlot called."            
    ' For Area Plot
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, COLOR)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)    
End Sub

Sub updateScatter(COLOR As Long)
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_EDGECOLOR, COLOR)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLOR, COLOR)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_COLORREPEAT, &H00000002&)
	ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SSA_SIZE, &H00000020&)
    ' fixme scattersize=0.032 Innches
    ' MsgBox "updateScatter called."        
End Sub

Sub updateSolid(COLOR As Long)
	ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
	ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, COLOR)
	ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
    ' MsgBox "updateSolid called."            
End Sub

Function DetectPlotType() As String
    On Error GoTo ErrorHandler
    Dim ObjectType As Variant
    Dim object_type As Variant
    object_type = ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).Plots(0).GetAttribute(SLA_TYPE, ObjectType)
    If object_type = False Then
        DetectPlotType = "Error: Failed to get object type."
        Exit Function
    End If
    ' Map the ObjectType to a string
    Select Case object_type
        Case SLA_TYPE_SCATTER, SLA_TYPE_MINVAL, SLA_TYPE_POLARXY, SLA_TYPE_3DBAR, SLA_TYPE_TERNARYSCATTER
            DetectPlotType = "Scatter/Line"
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
    DebugMsg "Type Detected: " & DetectPlotType
    Exit Function
ErrorHandler:
    DetectPlotType = "An error has occurred: " & Err.Description
End Function
Sub Main()
    On Error GoTo ErrorHandler

	Dim FullPATH As String
    Dim OrigPageName As String
    Dim ObjectType As String
    Dim COLOR As Long
    
    ' Remember the original page
    FullPATH = ActiveDocument.FullName
    OrigPageName = ActiveDocument.CurrentPageItem.Name
    ActiveDocument.NotebookItems(OrigPageName).IsCurrentBrowserEntry = True

    ' Get the color value for blue
    COLOR = getColor("Blue")
    
    ' Find the type of the object
    ObjectType = DetectObjectType()
    
    ' Check the object type and call the corresponding update function
    If ObjectType = "Scatter/Line/Area" Or ObjectType = "3D Scatter/Line" Then
        updatePlot COLOR
        updateScatter COLOR
    ElseIf ObjectType = "Bar" Or ObjectType = "Stacked Bar" Or ObjectType = "Box" Then
        updateSolid COLOR

    Else
        ' Raise a custom error
        Err.Raise vbObjectError + 513, "Main", "Unknown or unsupported object type: " & ObjectType
    End If
    
    ' Go back to the original page
	Notebooks(FullPATH).NotebookItems(OrigPageName).Open
	
    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred: " & Err.Description
End Sub
