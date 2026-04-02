VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "UserForm1"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents sh As ShapeHook
Attribute sh.VB_VarHelpID = -1


Private Sub btnTrackObjects_Click()
  Set sh = ShapeHook.CreateActive()
End Sub

Private Sub btnTrackingStop_Click()
  Set sh = Nothing
End Sub



Private Sub sh_ShapeAdded(ByVal shp As Shape)
  MsgBox "Shape Added " & shp.name
End Sub

Private Sub sh_ShapeMoved(ByVal shp As Shape, fromX As Double, fromY As Double, toX As Double, toY As Double)
  shp.Fill.ForeColor.RGB = RGB(Rnd() * 256, Rnd() * 256, Rnd() * 256)
End Sub

Private Sub sh_ShapeRemoved(ByVal name As String)
  MsgBox "Shape Removed " & name
End Sub

Private Sub sh_ShapeResized(ByVal shp As Shape, fromW As Double, fromH As Double, toW As Double, toH As Double)
  shp.Fill.BackColor.RGB = RGB(Rnd() * 256, Rnd() * 256, Rnd() * 256)
End Sub
