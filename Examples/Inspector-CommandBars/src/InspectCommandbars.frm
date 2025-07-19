VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InspectCommandbars 
   Caption         =   "CommandBar Inspector"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "InspectCommandbars.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InspectCommandbars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TThis
  CommandBarControls As stdArray
  QueriedControls As stdArray
  ListboxHeight As Long
End Type
Private This As TThis
Private WithEvents IC As MSINKAUTLib.InkCollector
Attribute IC.VB_VarHelpID = -1

'Userform initializer
'@private
Private Sub UserForm_Initialize()
  Set This.CommandBarControls = stdArray.Create()
  
  Set IC = New MSINKAUTLib.InkCollector
  With IC
    With stdWindow.CreateFromIUnknown(List)
      IC.hwnd = .Handle
    End With
    .SetEventInterest ICEI_MouseWheel, True
    .MousePointer = IMP_Arrow
    .DynamicRendering = False
    .DefaultDrawingAttributes.Transparency = 255
    .Enabled = True
  End With
  
  ModeSwitcher.AddItem "App"
  ModeSwitcher.AddItem "VBE"
  ModeSwitcher.value = "App"
  Call ModeSwitcher_Change
  
  Call UserForm_Resize
  
  stdWindow.CreateFromIUnknown(Me).isResizable = True
End Sub

'Helper for getting the selected CommandBarControl from the listbox
'@private
Private Property Get SelectedControl() As Dictionary
  Set SelectedControl = This.QueriedControls.item(List.ListIndex + 1)
End Property

'When modeswitcher combobox changes => Change query items
'@private
Private Sub ModeSwitcher_Change()
  Dim cbars As CommandBars, cmd As String
  If ModeSwitcher.value = "App" Then
    Set cbars = Application.CommandBars
    cmd = "Application.CommandBars(""$bar"").Controls(""$ctl"").Execute"
  Else
    Set cbars = Application.VBE.CommandBars
    cmd = "Application.VBE.CommandBars(""$bar"").Controls(""$ctl"").Execute"
  End If
  
  Set This.CommandBarControls = stdArray.Create()
  
  Dim cbar As CommandBar
  Dim index As Long
  For Each cbar In cbars
    Dim c As CommandBarControl
    For Each c In cbar.Controls
      'Create dictionary of key elements
      Dim o As Object: Set o = New Dictionary
      o("id") = c.ID
      o("parent") = cbar.name
      o("name") = c.Caption
      o("sanitizedName") = LCase(Replace(c.Caption, "&", ""))
      o("command") = Replace(Replace(cmd, "$bar", cbar.name), "$ctl", c.Caption)
      
      Set o("control") = c
      
      'Add to controls array
      Call This.CommandBarControls.Push(o)
    Next
  Next
  
  Set This.QueriedControls = This.CommandBarControls
  SearchBox.value = "" 'Clear searchbox
  Call UpdateListbox
End Sub

'When searchbox keyup => Search for controls
'@private
'@devNote - Some optimisation of searches is done here, when searching for queries contained within other queries
Private Sub SearchBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Dim searchText As String: searchText = LCase(SearchBox.value)
  If searchText = "" Then
    Set This.QueriedControls = This.CommandBarControls
  Else
    'Can we optimise the search?
    Dim optimisedScrape As Boolean
    Static lastSearch As String
    If lastSearch = searchText Then Exit Sub
    If searchText Like "*" & lastSearch & "*" And lastSearch <> "" Then
      optimisedScrape = True
    End If
    lastSearch = searchText
  
    
    Dim query As stdLambda: Set query = stdLambda.Create("$2.sanitizedName like $1").Bind("*" & searchText & "*")
    If optimisedScrape Then
      'If text is part of existing text then search within the already filtered controls
      Set This.QueriedControls = This.QueriedControls.Filter(query)
    Else
      Set This.QueriedControls = This.CommandBarControls.Filter(query)
    End If
  End If
  
  Call UpdateListbox
End Sub



'Update items within the listbox. This listbox has 3 columns showing the ID, Parent name and Control Caption
'@private
Private Sub UpdateListbox()
  List.Clear
  List.ColumnCount = 3
  List.ColumnWidths = "50,140,100"
  
  Dim control, index As Long: index = -1
  For Each control In This.QueriedControls
    index = index + 1
    List.AddItem "x", index
    List.Column(0, index) = control("id")
    List.Column(1, index) = control("parent")
    List.Column(2, index) = control("name")
  Next
  'Call This.QueriedControls.ForEach(stdLambda.Create("$1.AddItem($2.display)").Bind(List))
End Sub


'When userform resized => Change position and sizes of controls dynamically
'@private
'@eventHandler
Private Sub UserForm_Resize()
  SearchBox.width = Me.width - 35
  List.width = Me.width - 35
  List.height = Me.height - 110
  List.ColumnWidths = 50 / 323.5 * List.width & ";" & 140 / 323.5 * List.width & ";" & 100 / 323.5 * List.width
  btnPrint.top = Me.height - 53
  btnCopyMSO.top = Me.height - 53
  lblHelp.left = Me.width - lblHelp.width - 10
  lblHelp.top = Me.height - 53
  ModeSwitcher.left = Me.width - ModeSwitcher.width - 23
  This.ListboxHeight = stdAcc.CreateFromHwnd(IC.hwnd).Location("Height")
End Sub

'When button clicked => Copy MSO ID to clipboard.
'@private
'@eventHandler
Private Sub btnCopyMSO_Click()
  stdClipboard.Text = SelectedControl.item("id")
End Sub

'When button clicked => Print execute command to immediate window
'@private
'@eventHandler
Private Sub btnPrint_Click()
  Debug.Print SelectedControl.item("command")
End Sub

'Capture the scrollwheel => Scroll listbox view
'@private
'@eventHandler
'@param Button - the mouse button pressed (we don't care about this param)
'@param Shift - The direction of movement. Negative for down, postive for upward scroll.
Private Sub IC_MouseWheel(ByVal Button As MSINKAUTLib.InkMouseButton, ByVal Shift As MSINKAUTLib.InkShiftKeyModifierFlags, ByVal Delta As Long, ByVal x As Long, ByVal y As Long, Cancel As Boolean)
  Dim TargetProperty As String, CurrentValue As Long
  Static ItemHeight As Long: If ItemHeight = 0 Then ItemHeight = stdAcc.CreateFromHwnd(IC.hwnd).children(1).Location("Height")
  If Not Shift Then
    Dim iVisibleListBoxRowCount As Long: iVisibleListBoxRowCount = This.ListboxHeight / ItemHeight
    If Delta < 0 Then
      Delta = -1 * iVisibleListBoxRowCount
    Else
      Delta = iVisibleListBoxRowCount
    End If
    Dim newIndex As Long: newIndex = List.TopIndex - Delta
    If newIndex < 0 Then newIndex = 0
    If newIndex > This.QueriedControls.length - 1 Then newIndex = This.QueriedControls.length - 1
    List.TopIndex = newIndex
  End If
End Sub

'When listbox items are double clicked => execute item
'@private
'@eventHandler
'@param Cancel - Ignored param. Return true to cancel this event
Private Sub List_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  On Error Resume Next
  SelectedControl.item("control").Execute
End Sub

'Handler for ctrl+c => copy command to clipboard
'@private
'@eventHandler
Private Sub List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Const KEY_C As Long = 67, CTRL As Long = 2
  If KeyCode = KEY_C And Shift = CTRL Then
    stdClipboard.Text = SelectedControl.item("command")
  End If
End Sub
