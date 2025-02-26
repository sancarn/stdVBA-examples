VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Questionnaire 
   Caption         =   "Questionnaire"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   OleObjectBlob   =   "Questionnaire.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Questionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rows As stdEnumerator

Private Sub SubmitButton_Click()
  Dim template As String: template = shTemplate.Shapes("NoteTemplate").TextFrame2.TextRange.text
  Dim replacer As stdCallback: Set replacer = stdCallback.CreateFromObjectMethod(Me, "protReduceRow")
  stdClipboard.text = rows.reduce(replacer, template)
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  Dim lo As ListObject: Set lo = shTemplate.ListObjects("UserformElements")
  Set rows = stdEnumerator.CreateFromListObject(lo)
  
  Dim rowCreator As stdCallback: Set rowCreator = stdCallback.CreateFromObjectMethod(Me, "protCreateRow")
  Call rows.ForEach(rowCreator)
End Sub

Public Sub protCreateRow(ByVal row As Object, ByVal index As Long)
  Set row("ui") = CreateObject("Scripting.Dictionary")
  With row("ui")
    Set .Item("label") = stdUIElement.CreateFromType(Frame1.Controls, uiLabel, Caption:=row("Userform-Description"), fTop:=(index - 1) * 15, fWidth:=100)
    Dim element As stdUIElement
    Select Case row("Type")
      Case "Dropdown"
        Set element = stdUIElement.CreateFromType(Frame1.Controls, uiCombobox, fLeft:=100, fTop:=(index - 1) * 15, fWidth:=100)
        Dim cb As ComboBox: Set cb = element.uiObject
        cb.List = Split(row("dropdown-choices"), ";")
      Case "Checkbox"
        Set element = stdUIElement.CreateFromType(Frame1.Controls, uiCheckBox, fLeft:=100, fTop:=(index - 1) * 15)
    End Select
    Set .Item("input") = element
  End With
  
End Sub

Public Function protReduceRow(ByVal text As String, ByVal row As Object, ByVal index As Long) As String
  Dim finder As String: finder = "{" & row("TemplateName") & "}"
  Dim replacer As String
  Select Case row("Type")
    Case "Dropdown"
      replacer = row("ui")("input").Value
    Case "Checkbox"
      If row("ui")("input").Value Then
        replacer = row("checkbox-yes-text")
      Else
        replacer = row("checkbox-no-text")
      End If
  End Select
  protReduceRow = replace(text, finder, replacer)
End Function



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Unload Me
End Sub
