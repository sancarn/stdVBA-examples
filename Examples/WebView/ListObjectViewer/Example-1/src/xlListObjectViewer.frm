VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} xlListObjectViewer 
   Caption         =   "List Object Viewer"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14565
   OleObjectBlob   =   "xlListObjectViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "xlListObjectViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TThis
  w As stdWebView
  win As stdWindow
  lo As ListObject
End Type
Private This As TThis
Private WithEvents sht As Worksheet
Attribute sht.VB_VarHelpID = -1



Public Function Create(ByVal lo As ListObject, ByVal html As String) As xlListObjectViewer
  Set Create = New xlListObjectViewer
  Call Create.protInit(lo, html)
  Call Create.Show(False)
  
  Unload Me
End Function

Public Sub protInit(ByVal lo As ListObject, ByRef html As String)
  Set This.lo = lo
  Set sht = This.lo.parent
  Set This.w = stdWebView.CreateFromUserform(Me)
  Set This.win = stdWindow.CreateFromIUnknown(Me)
  This.w.html = html
  This.win.isAlwaysOnTop = True
  This.win.isResizable = True
  This.win.isAppWindow = True
  This.win.setOwnerHandle 0^
  This.win.isMaximiseButtonVisible = True
  This.win.isMinimiseButtonVisible = True
  This.win.isPopupWindow = True
  Call sht_SelectionChange(Selection)
End Sub

Public Property Get WebView() As stdWebView
  Set WebView = This.w
End Property


Private Sub sht_SelectionChange(ByVal Target As Range)
  If This.lo.DataBodyRange Is Nothing Then Exit Sub
  Dim rLo As Range: Set rLo = Application.Intersect(Target.EntireRow, This.lo.DataBodyRange)
  If rLo Is Nothing Then Exit Sub
  If rLo.Rows.CountLarge <> 1 Then Exit Sub
  
  Dim obj As Object: Set obj = CreateObject("Scripting.Dictionary")
  Dim header: header = This.lo.HeaderRowRange.value
  Dim values: values = rLo.value
  Dim j As Long
  For j = 1 To UBound(header, 2)
    obj(header(1, j)) = values(1, j)
  Next
  
  Call This.w.RemoveHostObject("listrow")
  Call This.w.AddHostObject("listrow", obj)
  On Error Resume Next
    Call This.w.JavaScriptRun("(function(){window.dispatchEvent(new CustomEvent('listrow-changed'));})();")
  On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
  This.w.Resize
End Sub

Private Sub UserForm_Terminate()
  Set sht = Nothing
  Set This.lo = Nothing
  Set This.w = Nothing
  Set This.win = Nothing
End Sub
