Attribute VB_Name = "MAnchoring"
'TODO: Replace with stdUIElement anchors after implementation, see https://github.com/sancarn/stdVBA/issues/112
Public Enum TElementAnchorType
  AnchorTypeFixed
  AnchorTypePercentile
End Enum
Public Type TElementAnchorPart
  Type As TElementAnchorType
  value As Double
End Type
Public Type TElementAnchor
  Parent As Object
  Element As Object
  Top As TElementAnchorPart
  Left As TElementAnchorPart
  Width As TElementAnchorPart
  Height As TElementAnchorPart
End Type

'@param Parent as Object<UserForm|Frame>
Public Function TElementAnchor_Create(ByVal Parent As Object, ByVal Element As Object, ByVal Top As TElementAnchorType, ByVal Left As TElementAnchorType, ByVal Width As TElementAnchorType, ByVal Height As TElementAnchorType) As TElementAnchor
  With TElementAnchor_Create
    Set .Parent = Parent
    Set .Element = Element
    .Top.Type = Top
    .Top.value = Element.Top / IIf(Top = AnchorTypeFixed, 1, Parent.InsideHeight)
    .Left.Type = Left
    .Left.value = Element.Left / IIf(Left = AnchorTypeFixed, 1, Parent.InsideWidth)
    .Width.Type = Width
    .Width.value = Element.Width / IIf(Width = AnchorTypeFixed, 1, Parent.InsideWidth)
    .Height.Type = Height
    .Height.value = Element.Height / IIf(Height = AnchorTypeFixed, 1, Parent.InsideHeight)
  End With
End Function

Public Sub TElementAnchor_Resize(ByRef anchor As TElementAnchor)
  With anchor
    If .Top.Type = AnchorTypePercentile Then
      .Element.Top = .Parent.InsideHeight * .Top.value
    End If
    If .Left.Type = AnchorTypePercentile Then
      .Element.Left = .Parent.InsideWidth * .Left.value
    End If
    If .Width.Type = AnchorTypePercentile Then
      .Element.Width = .Parent.InsideWidth * .Width.value
    End If
    If .Height.Type = AnchorTypePercentile Then
      .Element.Height = .Parent.InsideHeight * .Height.value
    End If
  End With
End Sub
