Attribute VB_Name = "modMain"
Sub ShowForm()
  With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Title = "Please select the JSON file to view"
    .Filters.Add "JSON", "*.json"
    If .Show Then
      Dim path As String: path = .SelectedItems(1)
      Call JSONViewer.ShowViewerFromFile(path)
    End If
  End With
End Sub

