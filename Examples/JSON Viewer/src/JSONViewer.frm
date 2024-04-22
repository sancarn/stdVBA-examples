VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JSONViewer 
   Caption         =   "Registry Viewer"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
   OleObjectBlob   =   "JSONViewer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JSONViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const UseDictionaryLateBinding = True
Private WithEvents tree As tvTree
Attribute tree.VB_VarHelpID = -1
Private SelectedEntry As Object


Public Sub ShowViewerFromString(ByVal data As String)
  Dim json As stdJSON: Set json = stdJSON.CreateFromString(data)
  Dim uf As JSONViewer: Set uf = New JSONViewer
  Call uf.ShowViewer(json)
End Sub


Public Sub ShowViewerFromFile(ByVal path As String)
  Dim json As stdJSON: Set json = stdJSON.CreateFromFile(path)
  Dim uf As JSONViewer: Set uf = New JSONViewer
  Call uf.ShowViewer(json)
End Sub

Public Sub ShowViewer(ByVal json As stdJSON)
  'Roots to render in tree
  Dim root As Object: Set root = CreateDictionary("key", "root", "value", json, "isJSON", True, "parent", Me)
  Dim roots As Collection: Set roots = New Collection: Call roots.Add(root)
  
  'Create tree
  Set tree = tvTree.Create( _
    JsonTree, _
    roots, _
    stdCallback.CreateFromObjectMethod(Me, "getItemID"), _
    stdCallback.CreateFromObjectMethod(Me, "getItemName"), _
    stdCallback.CreateFromObjectMethod(Me, "getItemChildren") _
  )
  
  ''Add context menu buttons
  ' With tree.ContextMenu
  '   Set btnCopyCode = .Controls.Add(msoControlButton, 1)
  '   With btnCopyCode
  '     .Caption = "&Copy stdVBA code"
  '     .FaceId = 19
  '     .Tag = "Copy"
  '   End With
  ' End With

  Call tree_OnSelected(roots(1))
  
  Call Me.Show
End Sub

'Private Sub btnCopyCode_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'  '...
'End Sub

'Get some unique item ID
Public Function getItemID(ByVal v As Object) As String
  If v("isJSON") Then
    getItemID = ObjPtr(v("value"))
  Else
    getItemID = ObjPtr(v("parent")) & "." & v("key")
  End If
End Function

'Get the name of the item to be displayed in the tree view
Public Function getItemName(ByVal v As Object) As String
  If v("isJSON") Then
    Select Case v("value").JsonType
      Case eJSONObject
        getItemName = simpleSerialise(v("key")) & " : Object"
      Case eJSONArray
        getItemName = v("key") & " : Array"
    End Select
  Else
    getItemName = simpleSerialise(v("key")) & ": " & simpleSerialise(v("value"))
  End If
End Function

'Obtain the children of the item
Public Function getItemChildren(ByVal v As Object) As Collection
  If v("isJSON") Then
    Set getItemChildren = v("value").ChildrenInfo
  Else
    Set getItemChildren = New Collection
  End If
End Function

Private Sub tree_OnRefresh(obj As Object)
  Call tree_OnSelected(obj)
End Sub

Private Sub tree_OnSelected(obj As Object)
  'Set selected item
  Set SelectedEntry = obj
End Sub

Private Function simpleSerialise(ByVal v As Variant) As String
  Select Case vartype(v)
    Case VbVarType.vbString
      simpleSerialise = """" & v & """"
    Case VbVarType.vbBoolean
      simpleSerialise = iif(v, "true", "false")
    Case VbVarType.vbNull
      simpleSerialise = "null"
    Case Else
      simpleSerialise = v
  End Select
End Function

Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Create a dictionary
'@returns - The dictionary
Private Function CreateDictionary(ParamArray children()) As Object
  #If UseDictionaryLateBinding Then
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
  #Else
    Set CreateDictionary = New Scripting.Dictionary
  #End If
  CreateDictionary.CompareMode = vbTextCompare

  Dim i As Long
  For i = LBound(children) To UBound(children) Step 2
    Call CreateDictionary.Add(children(i), children(i + 1))
  Next
End Function

