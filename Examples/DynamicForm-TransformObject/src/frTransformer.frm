VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frTransformer 
   Caption         =   "Transformer"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frTransformer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function AlterObject(ByVal obj As Object) As Object
  Dim ctrls As Collection: Set ctrls = New Collection
  Dim oForm As frTransformer: Set oForm = New frTransformer
  Dim index As Long: index = -2
  Dim prop
  For Each prop In stdCOM.Create(obj).Properties
    index = index + 2
    ctrls.add stdUIElement.CreateFromType(oForm.controls, uiLabel, prop & "_label", prop, fTop:=index * 20)
    ctrls.add stdUIElement.CreateFromType(oForm.controls, uiTextBox, prop & "_field", , stdCallback.CreateFromObjectProperty(obj, prop, VbGet)(), _
      stdLambda.Create("if $3 = EUIElementEvent.uiElementEventKeyUp then let $1." & prop & " = $2.value end").Bind(obj), fTop:=(index + 1) * 20)
  Next
  oForm.Show
  Set AlterObject = obj
End Function

