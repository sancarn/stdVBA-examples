VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegistryViewer 
   Caption         =   "Registry Viewer"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
   OleObjectBlob   =   "RegistryViewer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegistryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents tree As tvTree
Attribute tree.VB_VarHelpID = -1
Private WithEvents btnCopyCode As CommandBarButton
Attribute btnCopyCode.VB_VarHelpID = -1
Private SelectedEntry As stdReg

Private Sub btnCopyCode_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
  stdClipboard.text = "stdReg.CreateFromKey(""" & SelectedEntry.path & """)"
End Sub

Private Sub tree_OnRefresh(obj As Object)
  Call tree_OnSelected(obj)
End Sub

Private Sub tree_OnSelected(obj As Object)
  'Set selected item
  Set SelectedEntry = obj
  
  'Set address bar path to selkected item path
  RegistryAddress.value = SelectedEntry.path
  
  'Add items to listview
  RegistryItems.ListItems.Clear
  Dim item As stdReg
  For Each item In SelectedEntry.Items
    With RegistryItems.ListItems.Add(, , item.name)
      Call .ListSubItems.Add(, , ItemTypeText(item.ItemType))
      Call .ListSubItems.Add(, , item.value)
    End With
  Next
End Sub

'Get item type to text description
Private Function ItemTypeText(ByVal it As ERegistryValueType) As String
  Select Case it
    Case ERegistryValueType.Value_Binary:             ItemTypeText = "BINARY"
    Case ERegistryValueType.Value_DWORD:              ItemTypeText = "DWORD"
    Case ERegistryValueType.Value_DWORD_BE:           ItemTypeText = "DWORD_BE"
    Case ERegistryValueType.Value_Link:               ItemTypeText = "LINK"
    Case ERegistryValueType.Value_None:               ItemTypeText = "NONE"
    Case ERegistryValueType.Value_QWORD:              ItemTypeText = "QWORD"
    Case ERegistryValueType.Value_String:             ItemTypeText = "STRING"
    Case ERegistryValueType.Value_String_Array:       ItemTypeText = "STRING_ARRAY"
    Case ERegistryValueType.Value_String_WithEnvVars: ItemTypeText = "STRING_WITH_ENV"
  End Select
End Function

Private Sub UserForm_Initialize()
  'Roots to render in tree
  Dim roots As Collection: Set roots = New Collection
  Call roots.Add(stdReg.Create("HKEY_CURRENT_USER"))
  Call roots.Add(stdReg.Create("HKEY_LOCAL_MACHINE"))
  Call roots.Add(stdReg.Create("HKEY_CLASSES_ROOT"))
  Call roots.Add(stdReg.Create("HKEY_USERS"))
  
  'Create tree
  Set tree = tvTree.Create( _
    RegistryKeys, _
    roots, _
    stdLambda.Create("$1.Path"), _
    stdLambda.Create("$1.Name"), _
    stdLambda.Create("$1.Keys") _
  )
  '
  With tree.ContextMenu
    Set btnCopyCode = .Controls.Add(msoControlButton, 1)
    With btnCopyCode
      .Caption = "&Copy stdVBA code"
      .FaceId = 19
      .Tag = "Copy"
    End With
  End With
  
  'Add columns to listview
  RegistryItems.View = lvwReport
  RegistryItems.ColumnHeaders.Add text:="Name"
  RegistryItems.ColumnHeaders.Add text:="Type"
  RegistryItems.ColumnHeaders.Add text:="Data"
  
  'Select first element
  Call tree_OnSelected(roots(1))
End Sub


