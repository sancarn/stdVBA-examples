VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tvTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TThis
  tv As TreeView
  roots As Collection
  getID As stdICallable
  getName As stdICallable
  getChildren As stdICallable
  getIcon As stdICallable
  iconList As ImageList
  nodes As Object
  ContextMenu As CommandBar
End Type
Private This As TThis
Private WithEvents tv As TreeView
Attribute tv.VB_VarHelpID = -1
Private WithEvents btnRefresh As CommandBarButton
Attribute btnRefresh.VB_VarHelpID = -1
Public Event OnSelected(obj As Object)
Public Event OnRefresh(obj As Object)

Public Function Create(ByVal tv As TreeView, ByVal roots As Collection, ByVal getID As stdICallable, ByVal getName As stdICallable, ByVal getChildren As stdICallable, Optional ByVal getIcon As stdICallable, Optional ByVal iconList As ImageList) As tvTree
  Set Create = New tvTree
  Call Create.protInit(tv, roots, getID, getName, getChildren, getIcon, iconList)
End Function
Public Sub protInit(ByVal otv As TreeView, ByVal roots As Collection, ByVal getID As stdICallable, ByVal getName As stdICallable, ByVal getChildren As stdICallable, ByVal getIcon As stdICallable, ByVal iconList As ImageList)
  Set This.tv = otv
  Set This.roots = roots
  Set This.getID = getID
  Set This.getName = getName
  Set This.getChildren = getChildren
  Set This.getIcon = getIcon
  Set This.iconList = iconList
  Set This.nodes = CreateObject("Scripting.Dictionary")
  Set tv = otv
  Set This.ContextMenu = Application.CommandBars.Add(getGUID(), msoBarPopup)
  With This.ContextMenu
    Set btnRefresh = .Controls.Add(msoControlButton)
    With btnRefresh
      .Caption = "&Refresh"
      .FaceId = 1759
      .Tag = "Refresh"
    End With
  End With
  
  tv.LineStyle = tvwRootLines
  tv.Indentation = 0
  tv.LabelEdit = tvwManual
  tv.Font.size = 12
  
  Dim root As Object
  For Each root In roots
    Dim sKey As String: sKey = This.getID.Run(root)
    Set This.nodes(sKey) = root
    Dim iImage As Variant: If This.getIcon Is Nothing Then iImage = GetMissing Else iImage = This.getIcon.Run(root)
    tv.nodes.Add key:=sKey, text:=This.getName.Run(root), Image:=iImage
    tv.nodes.Add sKey, tvwChild, getGUID(), "Dummy"
  Next
End Sub

Public Property Get ContextMenu() As CommandBar
  Set ContextMenu = This.ContextMenu
End Property



Private Function GetMissing(Optional vMissing As Variant) As Variant
  GetMissing = vMissing
End Function

Private Sub btnRefresh_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
  'Close and
  This.tv.SelectedItem.Expanded = False
  This.tv.nodes.Add This.tv.SelectedItem.key, tvwChild, getGUID(), "Dummy"
  
  RaiseEvent OnRefresh(This.nodes(This.tv.SelectedItem.key))
End Sub

Private Sub tv_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
  Const BTN_RIGHT_CLICK = 2
  
  'If right click
  If Button = BTN_RIGHT_CLICK Then
    Call This.ContextMenu.ShowPopup
  End If
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
  RaiseEvent OnSelected(This.nodes(Node.key))
End Sub



Private Sub tv_Expand(ByVal Node As MSComctlLib.Node)
  'Remove existing children
  If Node.children > 0 Then
    'Get first child
    Dim oRef As Object: Set oRef = Node.child
    
    Dim i As Long
    For i = 1 To Node.children
      Dim oOldRef As Object: Set oOldRef = oRef
      Set oRef = oRef.Next
      Call This.tv.nodes.Remove(oOldRef.key)
    Next
  End If
  
  Dim obj As Object: Set obj = This.nodes(Node.key)
  Dim oChild As Object
  For Each oChild In This.getChildren.Run(obj)
    Dim sKey As String: sKey = This.getID.Run(oChild)
    Set This.nodes(sKey) = oChild
    Dim iImage As Variant: If This.getIcon Is Nothing Then iImage = GetMissing Else iImage = This.getIcon.Run(oChild)
    tv.nodes.Add relative:=This.getID.Run(obj), relationship:=tvwChild, key:=sKey, text:=This.getName.Run(oChild), Image:=iImage
    tv.nodes.Add sKey, tvwChild, getGUID(), "Dummy"
  Next
  tv.Refresh
End Sub

Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function


