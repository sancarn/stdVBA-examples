VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccHelper 
   Caption         =   "Accessibility Helper"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
   OleObjectBlob   =   "AccHelper.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AccHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'in a module
Option Explicit

Public Shown As Boolean

Public Enum EInspectState
  Permanent
  Temporary
End Enum

Private pInspectorState As EInspectState
Private pDisabledTime As Date
Private oInspected As stdAcc



Private Sub DoDefaultAction_Click()
  stdWindow.CreateFromHwnd(oInspected.hwnd).Activate
  oInspected.DoDefaultAction
End Sub

Private Sub SetValue_Click()
  stdWindow.CreateFromHwnd(oInspected.hwnd).Activate
  oInspected.value = InputBox("What value due you want to set it to?")
  Call UpdateFromInspected
End Sub

Private Sub UserForm_Initialize()
    Me.Show
    Shown = True
End Sub

Public Sub Watch()
  stdWindow.CreateFromIUnknown(Me).isTopmost = True
  While Shown
    If Inspecting.value Then
      Set oInspected = stdAcc.CreateFromMouse()
      Call UpdateFromInspected
    End If
    
    'Handle temporary inspector
    If pInspectorState = Temporary Then
      Dim c: c = Now()
      Dim timeLeft As Double: timeLeft = Round(Second(Application.Max(pDisabledTime - Now(), 0)), 1)
      If timeLeft = 0 Then
        Inspecting.value = False
        pInspectorState = Permanent
      End If
      SecondsLeft.Caption = timeLeft
    End If
    
    DoEvents
  Wend
End Sub

Public Sub UpdateFromInspected()
  On Error Resume Next
  With oInspected
    Me.crName.value = .name & "[P:" & .parent.name & "]"
    Me.crDefaultAction.value = .DefaultAction
    Me.crDescription.value = .Description
    Me.crRole.value = .Role
    Me.crStates.value = .States
    Me.crValue.value = .value
    Me.crLocation.value = "X: " & .Location!left & " Y: " & .Location!top & " W: " & .Location!width & " H: " & .Location!height
    Me.crHWND.value = .hwnd
    With stdWindow.CreateFromHwnd(.hwnd)
      If .Exists Then
        Me.crAppName.value = .ProcessName
        Me.crWindowClass.value = .Class
      End If
    End With
    On Error Resume Next
        'Me.crPath.value = ""
        'Me.crPath.value = .getPath()
    On Error GoTo 0
  End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Shown = False
    Unload Me
End Sub

Private Sub Enable5_Click()
  pInspectorState = Temporary
  pDisabledTime = Now() + TimeSerial(0, 0, 5)
  Inspecting.value = True
End Sub
Private Sub Inspecting_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  pInspectorState = Permanent
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

