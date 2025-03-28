Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const AXIS_X As Long = 0
Const AXIS_Y As Long = 1
Const LABEL_PTS_00 As Variant = "0"
Const LABEL_PTS_08 As String = "111"

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
   If DEBUG_MODE Then
      MsgBox msg, vbInformation, "Debug Info"
   End If
End Sub

Sub SetTitleSize()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_00)
End Sub

Sub RemoveTitle()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_FLAG_TITLESUNALIGNED, 0)
End Sub

Sub SetXLabelSize()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_X)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub

Sub SetYLabelSize()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_AXIS).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SLA_SELECTDIM, AXIS_Y)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETOBJECTATTR, STA_SIZE, LABEL_PTS_08)
End Sub

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   ' SetTitleSize()
   RemoveTitle()
   SetXLabelSize()
   SetXLabelSize()
On Error GoTo 0
End Sub
