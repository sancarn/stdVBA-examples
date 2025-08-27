VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type TThis
  win As stdWindow
End Type
Private this As TThis




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Call RenderOverRange(Target)
End Sub

'
Sub RenderOverRange(ByVal rng As Range)
    Dim r As Rect: r = ObjRect(rng)

    Set this.win = stdWindow.CreateHighlightRect( _
      r.Left, _
      r.Top, _
      r.Right - r.Left, _
      r.Bottom - r.Top _
    )
End Sub

'
'@source Jaafar Tribak - https://www.mrexcel.com/board/threads/highlighting-the-cell-under-the-mouse-pointer-gdi-based.1091151/
Private Function ObjRect(ByVal Obj As Range) As Rect
    Dim oPane  As Pane
    Set oPane = ThisWorkbook.windows(1).ActivePane
    
    With Obj
        ObjRect.Left = oPane.PointsToScreenPixelsX(.Left - 1)
        ObjRect.Top = oPane.PointsToScreenPixelsY(.Top - 1)
        ObjRect.Right = oPane.PointsToScreenPixelsX(.Left + .width)
        ObjRect.Bottom = oPane.PointsToScreenPixelsY(.Top + .height)
    End With
End Function
