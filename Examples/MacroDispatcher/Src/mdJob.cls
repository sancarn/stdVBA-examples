VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mdJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const DEBUG_MODE = False

#If VBA7 Then
   Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
   Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private Type TDependencies
  dirty As Boolean
  c As Collection
  e As stdEnumerator
End Type
Private Type TThis
  id As String
  Status As String
  path As String
  macro As String
  openInReadOnly As Boolean
  executedAt As Long
  app As Application
  appHandle As LongPtr
  appPID As Long
  appVBE As stdWindow
  deps As TDependencies
End Type
Private This As TThis
Public Metadata As Object

Public Function Create(ByVal sPath As String, ByVal sMacro As String, ByVal ReadOnly As Boolean) As mdJob
  Set Create = New mdJob
  Call Create.protInit(sPath, sMacro, ReadOnly)
End Function
Public Sub protInit(ByVal sPath As String, ByVal sMacro As String, ByVal openInReadOnly As Boolean)
  This.id = getGUID()
  This.Status = "Initialising"
  This.path = sPath
  This.macro = sMacro
  This.openInReadOnly = openInReadOnly
  Set This.deps.c = New Collection
  Set This.deps.e = stdEnumerator.CreateEmpty()
End Sub
Public Sub protAddDependency(ByVal dependency As mdJob)
  This.deps.dirty = True
  This.deps.c.add dependency
End Sub

Public Sub protStep()
  Dim sOldStatus As String: sOldStatus = This.Status
  Select Case This.Status
    Case "Initialising"
      'Ensure enumerator initialised
      If This.deps.dirty Then
        Set This.deps.e = stdEnumerator.CreateFromIEnumVariant(This.deps.c)
        This.deps.dirty = False
      End If
      
      'Check if dependency errored
      'TODO: Optional dependencies?
      If This.deps.e.checkAny(stdLambda.Create("$1.Status like ""Error*""")) Then
        This.Status = "Error - Dependency failed"
        Exit Sub
      End If
      
      If This.deps.e.checkAll(stdLambda.Create("$1.Status = ""Complete""")) Then
        This.Status = "Ready"
        Exit Sub
      End If
  
    'If all dependencies complete
    Case "Ready"
      'Create the application as seperate process
      Set This.app = New Application
      This.app.AskToUpdateLinks = False
      This.app.DisplayAlerts = False
      This.app.AutomationSecurity = msoAutomationSecurityLow
      This.appHandle = This.app.hwnd
      This.appPID = stdWindow.CreateFromHwnd(This.appHandle).ProcessID
      Set This.appVBE = stdWindow.CreateFromHwnd(This.app.VBE.MainWindow.hwnd)
      
      'Run the specified macro
      Call This.app.Workbooks.open(This.path, False, This.openInReadOnly)
      
      'Ensure window is not visible but
      #If DEBUG_MODE Then
        This.app.width = 0
        This.app.height = 0
        This.app.top = 0
        This.app.left = 0
        This.app.Visible = True
      #End If
      
      Call This.app.OnTime(Now(), "'" & This.path & "'!" & This.macro)
      
      This.executedAt = GetTickCount()
      This.Status = "Executed"
    Case "Executed"
      DoEvents
      'Wait at least half a second to ensure macro is running
      If Abs(GetTickCount() - This.executedAt) > 500 Then This.Status = "Running"
    Case "Running"
      Dim bFinished As Boolean: bFinished = False
      Dim wnd As stdWindow: Set wnd = getErrorWindow()
      If Not wnd Is Nothing Then
        This.Status = "Error - VBA Error: " & wnd.getAllText()
        If This.Status Like "Error - VBA Error: Microsoft Visual Basic*&End*" Then
          'Force quit doesn't work on these message boxes
          Call stdAcc.CreateFromHwnd(wnd.handle).FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""End""")).DoDefaultAction
        Else
          Call wnd.ForceQuit
        End If
        bFinished = True
      End If
      
      'Check if the macro has either errored or if it's finished
      If This.appVBE.Caption Like "*[[]Running[]]*" Or This.appVBE.Caption Like "*[[]Break[]]*" Then
        DoEvents
      ElseIf Not bFinished Then
        This.Status = "Complete"
        bFinished = True
      End If
      
      'Cleanup if finished
      If bFinished Then
        'Debug.Assert False
        This.app.Quit
      End If
  End Select
  If This.Status = "" Then Debug.Assert False
  
  'Report on status change
  If sOldStatus <> This.Status Then
    Debug.Print This.Status & " - '" & This.path & "'!" & This.macro
  End If
End Sub

Public Property Get Status() As String
  Status = This.Status
End Property

'Obtain VBA's error window. Useful for detecting errors raised by VBA engine which may be unpredictable.
'@returns {stdWindow} Obtain error window.
Public Function getErrorWindow() As stdWindow
  'Would use stdLambda but avoiding it for performance reasons
  Dim wnd As stdWindow, bFound As Boolean
  For Each wnd In stdWindow.CreateFromDesktop().children
    If wnd.Exists Then
      If wnd.ProcessID = This.appPID And wnd.Class = "#32770" Then
        bFound = True
        Exit For
      End If
    End If
  Next
  
  'Return window
  If wnd Is Nothing Then Exit Function
  If bFound And wnd.Exists Then Set getErrorWindow = wnd
End Function


















Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

