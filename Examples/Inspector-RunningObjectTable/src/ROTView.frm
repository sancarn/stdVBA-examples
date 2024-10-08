VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ROTView 
   Caption         =   "ROT Viewer"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915.001
   OleObjectBlob   =   "ROTView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ROTView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TThis
  monikers As Collection
  selectedMoniker As Object
  
  propView As uiFields
End Type
Private This As TThis

Private Sub btnRefresh_Click()
  Call RefreshMonikers
End Sub

Private Sub lbMonikers_Change()
  If lbMonikers.ListIndex >= 0 Then
    Set This.selectedMoniker = This.monikers(lbMonikers.ListIndex + 1)
    Call This.propView.UpdateSelection(This.selectedMoniker)
  End If
End Sub

Private Sub UserForm_Initialize()
  Set This.propView = uiFields.Create(frProps)
  With This.propView
    .keyWidthMultiplier = 0.25
    Call .AddField("Name", stdLambda.Create("$1.Name"))
    Call .AddField("Object Type", stdLambda.Create("$1.Type"))
    Call .AddField("ProgID", stdLambda.Create("$1.ProgID"))
  End With
  
  Call RefreshMonikers
End Sub

Private Sub RefreshMonikers()
  Set This.monikers = stdCOM.CreateFromActiveObjects()
  
  With lbMonikers
    .SetFocus
    .Clear
    
    Dim moniker
    For Each moniker In This.monikers
      Call .AddItem(moniker("Name"))
    Next
    
    .ListIndex = 0
  End With
End Sub


