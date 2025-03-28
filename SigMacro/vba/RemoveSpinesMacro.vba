Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const TICK_THICKNESS_00 As Variant = &H00000000

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
   If DEBUG_MODE Then
      MsgBox msg, vbInformation, "Debug Info"
   End If
End Sub

Sub RemoveXSpines()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_00)
End Sub

Sub RemoveYSpines()
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SUB2OPTIONS, TICK_THICKNESS_00)
End Sub

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   RemoveXSpines()
   RemoveYSpines()
On Error GoTo 0
End Sub
