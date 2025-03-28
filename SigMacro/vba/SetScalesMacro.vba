Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const WORKSHEET_NAME As String = "worksheet"
' Column indices for scale types
Const X_SCALE_TYPE_COL = 2
Const Y_SCALE_TYPE_COL = 8

' Axis indices
Const AXIS_X As Long = 0
Const AXIS_Y As Long = 1

' Axis scale types
Const SAA_TYPE_LINEAR = 1
Const SAA_TYPE_COMMON = 2
Const SAA_TYPE_LOG = 3
Const SAA_TYPE_PROBABILITY = 4
Const SAA_TYPE_PROBIT = 5
Const SAA_TYPE_LOGIT = 6
Const SAA_TYPE_CATEGORY = 7
Const SAA_TYPE_DATETIME = 8

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
   If DEBUG_MODE Then
      MsgBox msg, vbInformation, "Debug Info"
   End If
End Sub

Function SetAxisType(axisIndex As Long, scaleType As Long)
   Dim axis As Object
   ' Get the axis object directly
   Set axis = ActiveDocument.CurrentPageItem.GraphPages(0).Graphs(0).Axes(axisIndex)
   ' Set scale type
   axis.SetAttribute(SAA_TYPE, scaleType)
End Function

Function GetScaleType(cellValue As Variant) As Long
   Dim scaleType As Long

   ' Convert string or number to appropriate scale type constant
   Select Case CStr(LCase(cellValue))
      Case "linear", "1"
         scaleType = SAA_TYPE_LINEAR
      Case "common", "common log", "2"
         scaleType = SAA_TYPE_COMMON
      Case "log", "natural log", "3"
         scaleType = SAA_TYPE_LOG
      Case "probability", "4"
         scaleType = SAA_TYPE_PROBABILITY
      Case "probit", "5"
         scaleType = SAA_TYPE_PROBIT
      Case "logit", "6"
         scaleType = SAA_TYPE_LOGIT
      Case "category", "7"
         scaleType = SAA_TYPE_CATEGORY
      Case "datetime", "date", "time", "8"
         scaleType = SAA_TYPE_DATETIME
      Case Else
        ' Default to linear if unrecognized
        scaleType = SAA_TYPE_LINEAR
   End Select

   DebugMsg CStr(GetScaleType)
   GetScaleType = scaleType
  
End Function

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   Dim xScaleData As Variant, yScaleData As Variant
   Dim xScaleType As Long, yScaleType As Long

   ' Read scale types from worksheet
   xScaleData = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(X_SCALE_TYPE_COL, 0, X_SCALE_TYPE_COL, 0)
   yScaleData = ActiveDocument.NotebookItems(WORKSHEET_NAME).DataTable.GetData(Y_SCALE_TYPE_COL, 0, Y_SCALE_TYPE_COL, 0)

   ' Convert to scale type constants
   xScaleType = GetScaleType(xScaleData(0, 0))
   yScaleType = GetScaleType(yScaleData(0, 0))

   ' Set X axis scale type
   SetAxisType AXIS_X, xScaleType
   If Err.Number <> 0 Then
      DebugMsg "Error setting X axis scale: " & Err.Description
      Err.Clear
   End If

   ' Set Y axis scale type
   SetAxisType AXIS_Y, yScaleType
   If Err.Number <> 0 Then
      DebugMsg "Error setting Y axis scale: " & Err.Description
      Err.Clear
   End If

On Error GoTo 0
End Sub
