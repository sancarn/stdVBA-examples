VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Test 
   Caption         =   "Test"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Test.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private textBox As uiTextBoxEx

Private Sub MsgboxMe_Click()
  MsgBox textBox.Text, vbInformation
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  Set textBox = uiTextBoxEx.Create(Me.TextBoxFrame)
  
  With textBox
    .Text = "This is the UI Textbox - an " & _
            "example of stdVBA in action..."
    .ReadOnly = True
  End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Call textBox.Terminate
  
  'Have to terminate VBA here.
  End
End Sub


