Attribute VB_Name = "mCommands"
Sub ShowForm()
  AccessibilityInspector.Show
End Sub

Sub ShowForRoot(ByVal root As tvAcc)
  Call AccessibilityInspector.CreateForRoot(root)
End Sub

Sub dumpClassesAll()
  Call dumpClasses(stdShell.CreateFile("C:\Temp\classes.txt"), stdAcc.CreateFromDesktop())
End Sub

Sub dumpClasses(file As stdShell, ByVal acc As stdAcc)
  Dim wnd As stdWindow
  Set wnd = stdWindow.CreateFromHwnd(acc.hwnd)
  If wnd.Exists Then Call file.Append(wnd.Class & vbCrLf)
  Dim child As stdAcc
  For Each child In acc.children
    DoEvents
    Call dumpClasses(file, child)
  Next
End Sub
