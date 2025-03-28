Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const AXIS_X As Long = 1
Const AXIS_Y As Long = 2
Const HORIZONTAL_DIRECTION As Long = 1
Const VERTICAL_DIRECTION As Long = 2
Const TICK_THICKNESS_08 As Variant = &H00000008
Const TICK_LENGTH_20 As Variant = &H00000020

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
   If DEBUG_MODE Then
      MsgBox msg, vbInformation, "Debug Info"
   End If
End Sub

Sub SetXTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, HORIZONTAL_DIRECTION)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_20)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, VERTICAL_DIRECTION)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_20)    
End Sub

Sub SetYTickSizes()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, HORIZONTAL_DIRECTION)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_20)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_SELECTLINE, VERTICAL_DIRECTION)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SEA_THICKNESS, TICK_THICKNESS_08)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETAXISATTR, SAA_TICSIZE, TICK_LENGTH_20)    
End Sub

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   SetXTickSizes()
   SetYTickSizes()
   SetXTickSizes()
   SetYTickSizes()   
On Error GoTo 0
End Sub