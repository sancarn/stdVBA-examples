Attribute VB_Name = "Logger"
Private shell As stdShell
Sub GlobalLog(ByVal Class As String, ByVal message As String)
  If shell Is Nothing Then Set shell = stdShell.CreateFile("C:\Temp\XLClipboardLog.txt")
  Dim sMsg As String: sMsg = Class & " -- " & message
  Call shell.Append(sMsg)
  Debug.Print sMsg
  
End Sub
