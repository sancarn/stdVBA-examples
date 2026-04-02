Attribute VB_Name = "mShowForm"
Private viewer As xlListObjectViewer
Sub ShowForm()
  Dim htmlFile As String: htmlFile = ThisWorkbook.path & Application.PathSeparator & "index.html"
  Dim htmlText As String: htmlText = stdShell.Create(htmlFile).ReadText()
  Dim lo As ListObject: Set lo = shEmployees.ListObjects("Employees")
  Set viewer = xlListObjectViewer.Create(lo, htmlText)
End Sub
