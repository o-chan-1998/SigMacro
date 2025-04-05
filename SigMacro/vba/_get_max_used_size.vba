Dim X As Long 
Dim Y As Long 
ActiveDocument.NotebookItems(2).DataTable.GetMaxUsedSize(X,Y) 
MsgBox CStr(X) + ", " + CStr(Y)  