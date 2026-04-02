Attribute VB_Name = "udfs"
Public Function udfBuildJSONObject(ParamArray params()) As String
  Dim i As Long
  With stdJSON.Create(eJSONObject)
    For i = 0 To UBound(params) Step 2
      .Add params(i), params(i + 1)
      udfBuildJSONObject = .ToString()
    Next
  End With
End Function
