Option Explicit

' ========================================
' Constants
' ========================================
Const DEBUG_MODE As Boolean = True
Const TITLE_HIDE As Long = 0

' ========================================
' Functions
' ========================================
Sub DebugMsg(msg As String)
    If DEBUG_MODE Then
    MsgBox msg, vbInformation, "Debug Info"
    End If
End Sub

Sub RemoveTitle()
    ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_FLAG_TITLESUNALIGNED, TITLE_HIDE)
    ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETGRAPHATTR, SGA_TITLESHOW, TITLE_HIDE)
End Sub

' ========================================
' Main
' ========================================
Sub Main()
   On Error Resume Next
   RemoveTitle()
   On Error GoTo 0
End Sub
