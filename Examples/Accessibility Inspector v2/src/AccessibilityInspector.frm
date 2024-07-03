VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccessibilityInspector 
   Caption         =   "Accessibility Inspector"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
   OleObjectBlob   =   "AccessibilityInspector.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AccessibilityInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'[P] - Planned
'[?] - Maybe
'[ ] - Not planned
'
'Features:
'|- Toolbar
'|  |- [P] Refresh
'|  |- [ ] Watch Focus
'|  |- [ ] Watch Caret
'|  |- [P] Watch Cursor
'|  |- [P] Watch Cursor 5s
'|  |- [?] Watch Tooltip
'|  |- [P] Show Highlight Rectangle
'|  |- [?] Show Tooltip
'|  |- [P] Focus
'|  |- [P] Select
'|  |- [P] Show Context Menu
'|- Properties
'|  |- [P] Identity
'|  |- [P] Name
'|  |- [P] Value              [SetValue]
'|  |- [P] Default Action     [DoDefaultAction]
'|  |- [P] Description
'|  |- [P] Role
'|  |- [P] States
'|  |- [P] Location
'|  |- [P] HWND
'|  |- [P] Program
'|  |- [P] WND Class
'|  |- [P] Path


Private Type TInit
  AllTop As Double
  TreeControlLeft As Double
  TreeControlWidth As Double
  FieldsLeft As Double
  FieldsWidth As Double
  
  
  'As percentages of width
  pcTCLeft As Double
  pcTCWidth As Double
  pcFdLeft As Double
  pcFdWidth As Double
  pcAllHeight As Double
End Type
Private Type TWatch
  OldControl As stdAcc
  DateStarted As Date
End Type
Private Type TWatch5
  DateStarted As Date
  secDiff As Long
End Type
Private Type TThis
  init As TInit
  props As uiFields
  SelectedElement As tvAcc
  ProcWatch As TWatch
  ProcWatch5 As TWatch5
  HighlightRect As stdWindow
End Type
Private This As TThis

Private WithEvents tree As tvTree
Attribute tree.VB_VarHelpID = -1
Private WithEvents btnCopyCode As CommandBarButton
Attribute btnCopyCode.VB_VarHelpID = -1





Private Sub tree_OnRefresh(obj As Object)
  Call tree_OnSelected(obj)
End Sub

Private Sub tree_OnSelected(obj As Object)
  'Set selected item
  Set This.SelectedElement = obj
  Call This.props.UpdateSelection(obj)
  
  If TypeOf obj Is tvAcc Then
    Dim t As tvAcc: Set t = obj
    If btnHighlightRectangles.SpecialEffect = fmSpecialEffectSunken Then
      Dim loc: Set loc = t.Location
      Set This.HighlightRect = stdWindow.CreateHighlightRect(loc("left"), loc("top"), loc("width"), loc("height"), 10, RGB(255, 255, 0))
      DoEvents
    End If
  End If
  
End Sub

Private Sub btnSearch_Click()
  btnSearch.SpecialEffect = fmSpecialEffectSunken
  Me.Hide
  textToSearch = "*" & LCase(InputBox("What do you want to search for?")) & "*"
  Me.Show
  Dim n As Node
  For Each n In TreeControl.nodes
    If LCase(n.Text) Like textToSearch Then
      n.EnsureVisible
      n.BackColor = RGB(255, 255, 0)
    Else
      n.BackColor = RGB(255, 255, 255)
    End If
  Next
  btnSearch.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub UserForm_Initialize()
  With This.init
    .AllTop = TreeControl.top
    .TreeControlLeft = TreeControl.left
    .TreeControlWidth = TreeControl.width
    .FieldsLeft = PropertyFrame.left
    .FieldsWidth = PropertyFrame.width
    
    Dim FormWidth As Double: FormWidth = Me.width
    Dim FormHeight As Double: FormHeight = Me.height
    
    .pcTCLeft = .TreeControlLeft / FormWidth
    .pcTCWidth = .TreeControlWidth / FormWidth
    .pcFdLeft = .FieldsLeft / FormWidth
    .pcFdWidth = .FieldsWidth / FormWidth
    .pcAllHeight = TreeControl.height / FormHeight
  End With
  With stdWindow.CreateFromIUnknown(Me)
    .isResizable = True 'set resizable
    .isAlwaysOnTop = True
    Call .setOwnerHandle(0)
  End With
  
  
  'Roots to render in tree
  Dim roots As Collection: Set roots = New Collection
  Call roots.Add(tvAcc.CreateFromDesktop())
  
  'Create tree
  Set tree = tvTree.Create( _
    TreeControl, _
    roots, _
    stdLambda.Create("$1.Identity"), _
    stdCallback.CreateFromObjectMethod(Me, "accFilter"), _
    stdLambda.Create("""'"" & $1.Name & ""' - "" & mid($1.Role,6)"), _
    stdLambda.Create("$1.Children") _
  )
  
  Set This.props = uiFields.Create(PropertyFrame)
  With This.props
    Call .AddField("Identity", stdLambda.Create("$1.Identity"))
    Call .AddField("Name", stdLambda.Create("$1.Name"))
    Call .AddField("Description", stdLambda.Create("$1.Description"))
    Call .AddField("Value", stdLambda.Create("$1.Value"), stdCallback.CreateFromObjectMethod(Me, "ElementSetValue"))
    Call .AddField("Default Action", stdLambda.Create("$1.DefaultAction"), stdLambda.Create("$1.DoDefaultAction"))
    Call .AddField("Role", stdLambda.Create("$1.Role"))
    Call .AddField("States", stdLambda.Create("$1.States"))
    Call .AddField("Location", stdCallback.CreateFromObjectMethod(Me, "getElementLocation"))
    Call .AddField("Hwnd", stdLambda.Create("$1.hwnd"))
    Call .AddField("Program", stdCallback.CreateFromObjectMethod(Me, "getApplicationPath"))
    Call .AddField("Window Class", stdCallback.CreateFromObjectMethod(Me, "getWindowClass"))
    Call .AddField("Path", stdLambda.Create("$1.getPath()"))
  End With
  
  'Select first element
  Call tree_OnSelected(roots(1))
End Sub

'Returns false if element should not be shown, else return true
Public Function accFilter(ByVal acc As tvAcc) As Boolean
  ''Debugging...
  'With acc
  '  Debug.Print "Acc: " & Join(Array(.Role, .States, .Children.Count, .hwnd), ";")
  '  With stdWindow.CreateFromHwnd(.hwnd)
  '    Debug.Print "Window: " & Join(Array(.Caption, .Visible, .handle, .Class, .ProcessName), ";")
  '    With stdProcess.CreateFromProcessId(.ProcessID)
  '      Debug.Print "Process: " & Join(Array(.name, .path, .Priority, .isRunning), ";")
  '    End With
  '  End With
  'End With
  
  DoEvents
  
  With acc
    If btnVisibleOnly.SpecialEffect = fmSpecialEffectSunken Then
      If (.StateData And STATE_INVISIBLE) = STATE_INVISIBLE Then Exit Function
    End If
    Select Case CLng(.hwnd)
      Case 66590, 66592: Exit Function
    End Select
    With stdWindow.CreateFromHwnd(.hwnd)
      
      If Not .Exists Then Exit Function
      If Not .Visible Then Exit Function
      With stdProcess.CreateFromProcessId(.ProcessID)
        If .name Like "CodeSetup-stable-*.tmp" Then
          Debug.Print acc.Role
          Debug.Print .name
          Debug.Print .path
          Exit Function
        End If
      End With
    End With
    'If .Identity = "Unknown" Then Exit Function
  End With
  
  accFilter = True                              'Show element
End Function
Public Function getApplicationPath(ByVal acc As tvAcc) As String
  With stdWindow.CreateFromHwnd(acc.hwnd)
    If Not .Exists Then Exit Function
    Dim pid As Long: pid = .ProcessID
    With stdProcess.CreateFromProcessId(.ProcessID)
      getApplicationPath = .path
    End With
  End With
End Function
Public Function getWindowClass(ByVal acc As tvAcc) As String
  With stdWindow.CreateFromHwnd(acc.hwnd)
    If .Exists Then
      getWindowClass = .Class
    Else
      getWindowClass = "Window no longer exists..."
    End If
  End With
End Function
Public Function getElementLocation(ByVal acc As tvAcc) As String
  If acc.Location Is Nothing Then Exit Function
  getElementLocation = "X: " & acc.Location!left & " Y: " & acc.Location!top & " W: " & acc.Location!width & " H: " & acc.Location!height
End Function
Public Sub ElementSetValue(ByVal acc As tvAcc)
  Me.Hide
  acc.value = InputBox("Enter the value to input")
  Me.Show
End Sub

Private Sub FollowMouse()
  Dim Mouse5Initialised As Boolean
  While btnFollowMouse.SpecialEffect = fmSpecialEffectSunken Or btnFollowMouse5.SpecialEffect = fmSpecialEffectSunken
    DoEvents
    
    'Trigger Follow mouse operations
    If btnFollowMouse.SpecialEffect = fmSpecialEffectSunken Then
'      Dim tMAcc As stdAcc: Set tMAcc = stdAcc.CreateFromMouse()
'      If This.ProcWatch.OldControl Is Nothing Then
'        Set This.ProcWatch.OldControl = tMAcc
'        This.ProcWatch.DateStarted = Now()
'      Else
'        If This.ProcWatch.OldControl.Identity <> tMAcc.Identity Then
'          Set This.ProcWatch.OldControl = tMAcc
'          This.ProcWatch.DateStarted = Now()
'        End If
'      End If
'
'      'If more than 2s has expired, set current control
'      If DateDiff("s", This.ProcWatch.DateStarted, Now()) >= 2 Then
'        Call tree_OnSelected(This.ProcWatch.OldControl)
'      End If
      Call tree_OnSelected(tvAcc.CreateFromMouse())
    End If
    
    'Trigger Follow Mouse 5 operations
    If btnFollowMouse5.SpecialEffect = fmSpecialEffectSunken Then
      Call tree_OnSelected(tvAcc.CreateFromMouse())
      Dim tM5SecDiff As Long: tM5SecDiff = DateDiff("s", This.ProcWatch5.DateStarted, Now())
      If Mouse5Initialised Then
        If tM5SecDiff <> This.ProcWatch5.secDiff Then
          If tM5SecDiff >= 5 Then
            btnFollowMouse5.SpecialEffect = fmSpecialEffectRaised
            Me.Caption = "Accessibility Inspector"
          Else
            This.ProcWatch5.secDiff = tM5SecDiff
            Me.Caption = "Accessibility Inspector - Countdown: " & (5 - tM5SecDiff)
          End If
        End If
      Else
        Mouse5Initialised = True
        This.ProcWatch5.secDiff = tM5SecDiff
        Me.Caption = "Accessibility Inspector - Countdown: " & (5 - tM5SecDiff)
      End If
      
        
      
    End If
  Wend
End Sub

Private Function getSelector(acc As tvAcc, Optional depth As Long = 0) As String
  Dim command As String
  If depth = 0 Or (x) Then
    If acc.name <> "" Then
      command = "acc.FindFirst(stdLambda.Create(""$1.Name = """"$NAME"""" and $1.Role = """"$ROLE""""""))"
      command = Replace(command, "$NAME", acc.name)
      command = Replace(command, "$ROLE", acc.Role)
      GoTo Finish
    End If
  End If
  Select Case acc.Role
    Case "ROLE_PUSHBUTTON", "ROLE_RADIOBUTTON", "ROLE_LISTITEM", "ROLE_LINK"
      If acc.value <> "" Then
        command = "acc.FindFirst(stdLambda.Create(""$1.Value = """"$VALUE"""" and $1.Role = """"$ROLE""""""))"
        command = Replace(command, "$VALUE", acc.value)
        command = Replace(command, "$ROLE", acc.Role)
        GoTo Finish
      End If
    Case Else
      If acc.children.count > 0 Then
        Dim child As stdAcc
        For Each child In acc.children
          command = getSelector(child, depth + 1)
          If command <> "" Then
            command = command & ".parent"
            GoTo Finish
          End If
        Next
      End If
  End Select
  Exit Function
Finish:
  getSelector = command
End Function

Private Sub btnCopyVB_Click()
  Call toggleButtonState(btnCopyVB)
  Dim s As String: s = getSelector(This.SelectedElement)
  If s = "" Then
    MsgBox "Unable to identify selector", vbExclamation
  Else
    stdClipboard.Text = s
  End If
  Call toggleButtonState(btnCopyVB)
End Sub

Private Sub btnFollowMouse_Click()
  With btnFollowMouse
    Select Case .SpecialEffect
      Case fmSpecialEffectRaised
        .SpecialEffect = fmSpecialEffectSunken
        btnFollowMouse5.SpecialEffect = fmSpecialEffectRaised
        Call FollowMouse
      Case fmSpecialEffectSunken
        .SpecialEffect = fmSpecialEffectRaised
    End Select
  End With
End Sub

Private Sub btnFollowMouse5_Click()
  With btnFollowMouse5
    Select Case .SpecialEffect
      Case fmSpecialEffectRaised
        .SpecialEffect = fmSpecialEffectSunken
        btnFollowMouse.SpecialEffect = fmSpecialEffectRaised
        This.ProcWatch5.DateStarted = Now()
        Call FollowMouse
      Case fmSpecialEffectSunken
        .SpecialEffect = fmSpecialEffectRaised
    End Select
  End With
End Sub

Private Sub toggleButtonState(ByVal btn As Object)
  With btn
    Select Case .SpecialEffect
      Case fmSpecialEffectRaised
        .SpecialEffect = fmSpecialEffectSunken
      Case fmSpecialEffectSunken
        .SpecialEffect = fmSpecialEffectRaised
    End Select
  End With
End Sub

Private Sub btnHighlightRectangles_Click(): Call toggleButtonState(btnHighlightRectangles): End Sub
Private Sub btnVisibleOnly_Click(): Call toggleButtonState(btnVisibleOnly): End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Explicitely set to nothing, required else crash will occur
  Set This.HighlightRect = Nothing
End Sub

Private Sub UserForm_Resize()
  Dim width As Double: width = Me.width
  Dim height As Double: height = Me.height
  TreeControl.left = This.init.pcTCLeft * width
  TreeControl.width = This.init.pcTCWidth * width
  TreeControl.height = This.init.pcAllHeight * height
  This.props.left = This.init.pcFdLeft * width
  This.props.width = This.init.pcFdWidth * width
  This.props.height = This.init.pcAllHeight * height
End Sub

Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function
