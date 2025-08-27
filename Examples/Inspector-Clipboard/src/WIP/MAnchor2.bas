Attribute VB_Name = "MAnchor2"
'TODO: Replace with stdUIElement anchors after implementation, see https://github.com/sancarn/stdVBA/issues/112
Public Type TElementAnchorPart2
  Anchored As Boolean
  Initial As Double
End Type
Public Type TElementAnchor2
  Parent As Object
  Element As Object
  Top As TElementAnchorPart2
  Left As TElementAnchorPart2
  Bottom As TElementAnchorPart2
  Right As TElementAnchorPart2
End Type

'@param Parent as Object<UserForm|Frame>
Public Function TElementAnchor2_Create(ByVal Parent As Object, ByVal Element As Object, Optional ByVal Top As Boolean = True, Optional ByVal Left As Boolean = True, Optional ByVal Bottom As Boolean = True, Optional ByVal Right As Boolean = True) As TElementAnchor2
  With TElementAnchor2_Create
    Set .Parent = Parent
    Set .Element = Element
    .Top.Anchored = Top
    .Top.Initial = Element.Top
    .Left.Anchored = Left
    .Left.Initial = Element.Left
    .Bottom.Anchored = Bottom
    .Bottom.Initial = Element.Top + Element.Height
    .Right.Anchored = Right
    .Right.Initial = Element.Left + Element.Width
  End With
End Function

Public Sub TElementAnchor2_Resize(ByRef anchor As TElementAnchor2)
  On Error Resume Next
  With anchor
    ' Horizontal adjustment
    .Element.Left = IIf(.Left.Anchored, .Left.Initial, .Parent.Width - .Right.Initial - .Element.Width)
    .Element.Width = IIf(.Left.Anchored And .Right.Anchored, .Parent.Width - .Right.Initial, .Element.Width)
    
    ' Vertical adjustment
    .Element.Top = IIf(.Top.Anchored, .Top.Initial, .Parent.Height - .Bottom.Initial - .Element.Height)
    .Element.Height = IIf(.Top.Anchored And .Bottom.Anchored, .Parent.Height - .Top.Initial - .Bottom.Initial, .Element.Height)
  End With
End Sub
