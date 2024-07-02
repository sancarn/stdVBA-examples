VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements uiIMessagable
Private Type TFieldControl
  id As Long
  name As String
  getValue As stdICallable
  execute As stdICallable
  
  ctrlName As uiElement
  ctrlValue As uiElement
  ctrlExec As uiElement
End Type
Private Type TThis
  indexLookup As Object
  selected As Object
  uiFrame As MSForms.Frame
  fields() As TFieldControl
End Type
Private This As TThis

Public Function Create(ByVal fr As MSForms.Frame) As uiFields
  Set Create = New uiFields
  Call Create.protInit(fr)
End Function

Friend Sub protInit(ByVal fr As MSForms.Frame)
  Set This.uiFrame = fr
  Set This.indexLookup = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get Top() As Long
  Top = This.uiFrame.Top
End Property
Public Property Let Top(v As Long)
  This.uiFrame.Top = v
End Property
Public Property Get Left() As Long
  Left = This.uiFrame.Left
End Property
Public Property Let Left(v As Long)
  This.uiFrame.Left = v
End Property
Public Property Get width() As Long
  width = This.uiFrame.width
End Property
Public Property Let width(v As Long)
  This.uiFrame.width = v
  Dim iWidth As Long: iWidth = (This.uiFrame.InsideWidth - 4) / 2
  Dim i As Long: For i = 0 To UBound(This.fields)
    With This.fields(i)
      .ctrlName.width = iWidth
      .ctrlValue.Left = 2 + iWidth
      .ctrlValue.width = iWidth
      If Not .execute Is Nothing Then
        .ctrlValue.width = .ctrlValue.width - 10
        .ctrlExec.Left = 2 + 2 * iWidth - 10
      End If
    End With
  Next
End Property
Public Property Get height() As Long
  height = This.uiFrame.height
End Property
Public Property Let height(v As Long)
  This.uiFrame.height = v
End Property

Public Function AddField(ByVal sName As String, ByVal getValue As stdICallable, Optional ByVal execute As stdICallable = Nothing, Optional ByVal execIcon As StdPicture) As Long
  Dim iWidth As Long: iWidth = This.uiFrame.InsideWidth - 4
  On Error Resume Next
  Dim i As Long: i = UBound(This.fields) + 1
  On Error GoTo 0
  ReDim Preserve This.fields(0 To i)
  With This.fields(i)
    .id = i
    Set .ctrlName = uiElement.CreateFromType(Me, This.uiFrame.Controls, uiLabel, "Name_" & i, sName, 2, 2 + i * (20 + 2), iWidth / 2, 20)
    Set .ctrlValue = uiElement.CreateFromType(Me, This.uiFrame.Controls, uiTextBox, "Value_" & i, "", 2 + iWidth / 2, 2 + i * (20 + 2), iWidth / 2, 20)
    Set .getValue = getValue
    If Not execute Is Nothing Then
      .ctrlValue.width = .ctrlValue.width - 10
      Set .ctrlExec = uiElement.CreateFromType(Me, This.uiFrame.Controls, uiImage, "Button_" & i, "", 2 + iWidth - 10, 2 + i * (20 + 2), 10, 20)
      Set .execute = execute
      This.indexLookup(.ctrlExec.id) = i
    End If
    This.indexLookup(.ctrlName.id) = i
    This.indexLookup(.ctrlValue.id) = i
  End With
End Function


Public Sub UpdateSelection(ByVal obj As Object)
  Set This.selected = obj
  
  Dim i As Long
  For i = 0 To UBound(This.fields)
    This.fields(i).ctrlValue.value = This.fields(i).getValue.Run(obj)
  Next
End Sub

















Private Sub Class_Terminate()
  Set This.selected = Nothing
End Sub

Private Function uiIMessagable_Message(ByVal Caller As Object, iMessage As Long, Optional params As Variant = Empty) As Variant
  Dim el As uiElement: Set el = Caller
  If iMessage = EUIElementMessage.uiElementEventMouseClick Then
    Dim index As Long: index = This.indexLookup(el.id)
    With This.fields(index)
      If Not .execute Is Nothing Then Call .execute.Run(This.selected)
    End With
  End If
End Function
