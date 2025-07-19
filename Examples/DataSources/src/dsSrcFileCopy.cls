VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcFileCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements dsISrc

Private Function dsISrc_getName() As String
  dsISrc_getName = "FileCopy"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Set dsISrc_linkFiber = template
  With template
    .Meta("Source") = Data.item("Source")
    .Meta("Destination") = destPath
    
    Dim cb As stdCallback: Set cb = stdCallback.CreateFromObjectMethod(Me, "protProcessCopySingleFile").Bind(srcPath, destPath, True)
    Call .add(cb.Bind(1), "1. launching copy")
    Call .add(cb.Bind(2), "2. awaiting copy completion")
  End With
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("Source")
      .add "DisplayText", "Source File"
      .add "Tooltip", "The location of the file to copy."
      .add "Type", "File"
      With .AddObject("File")
        .add "Extensions", "*.*"
      End With
    End With
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function

'Processes an async GISSTdb OLEdb query.
'@fiberRunner
'@protected
Public Function protProcessCopySingleFile(ByVal State As Long, ByVal src As String, ByVal dest As String, ByVal required As Boolean, ByVal fiber As stdFiber) As Boolean
  Select Case State
    Case 1
      'If it exists, copy it whatever
      If FileExists(src) Then
        'Copy asynchronously
        Call shell("copy """ & src & """ """ & dest & """", vbHide)
      'If it doesn't exist but is required, error
      ElseIf required Then
        Call fiber.RaiseCriticalError("File '" & src & "' is not present but is required.")
        Exit Function
      End If
      protProcessCopySingleFile = True
    Case 2 ' await copy completion
      If FileExists(dest) Then protProcessCopySingleFile = FileLen(src) = FileLen(dest)
  End Select
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = (Dir(FilePath) <> "")
End Function


