Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const HIDE_LEGEND As Long = 0

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
   If DEBUG_MODE Then
      MsgBox msg, vbInformation, "Debug Info"
   End If
End Sub

Sub RemoveLegend()
   ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
   ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_AUTOLEGENDSHOW, HIDE_LEGEND)
End Sub

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   RemoveLegend()
On Error GoTo 0
End Sub
