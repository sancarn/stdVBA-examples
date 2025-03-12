VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sapSAPECC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Requires:
'* stdICallable
'* stdAcc
'* stdWindow
'* stdClipboard

'SAP shortcuts
'{F9} - Jump to first form value
'F1 - Documentation
'F2
'F3 - Back button  (Can exit to home, or jump back, disables on home page)
'F4 - Find objects in Classes (probably just a IH06 thing)
'F5
'F6
'F7
'F8 - Execute search
'F9 - Background print parameters (might differ between forms)
'F10
'F11 - Jump to Save as variant mmenu
'F12 - Cancel ( Can exit to home, or jump back, disables on home page)

'+F1 - Programmer documentation
'+F3 - Exit    Can exit to home OR Log off (if at home already)
'
'+F5 - Jump to Go to... Variant menu
'+F10 - Right click on current focussed item

'Alt F12 Open "Customise Local Layout" menu i.e. menu with abap debugger in

'USING SLEEP FOR CONSISTENCY
'--------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetTickCount Lib "Kernel32" () As Long
#Else
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetTickCount Lib "Kernel32" () As Long
#End If
#Const OPTION_INCLUDE_ASYNC = False
'--------------------------------------------------------------------------------------------------

'Workbook containing SAP data to close on cleanup
Private Type TThis
  wb As Workbook
  SAPMainWindow As stdWindow
  SAPMainAcc As stdAcc
  SAPToolbarWindow As stdWindow
  SAPToolbarAcc As stdAcc
  SAPTransactionAcc As stdAcc
  SAPErrBarAcc As stdAcc
  
  SAPChildWindow As stdWindow 'It appears that SAP uses the same child window for all operations, so save it after first requirement
  
  isTemporaryWindow As Boolean
End Type
Private This As TThis
Private Const Edge_Exe As String = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
Private Const ECC_URL As String = "http://sapportal.xxxxxxx/" 'URL to navigate to in order to launch SAP ECC


Public Function Create(Optional ByVal closeOnDestroy = True) As sapSAPECC
  Set Create = New sapSAPECC
  
  'Get Existing windows
  Dim wins As Collection: Set wins = getSapWindows()
  
  'Create new window via link
  Dim EdgeProc As stdProcess: Set EdgeProc = stdProcess.Create("""" & Edge_Exe & """ --new-window """ & ECC_URL & """")
  
  'Await for new window
  While getSapWindows(wins).count = 0
    DoEvents
  Wend
  Set wins = getSapWindows(wins)
  Dim sapWindow As stdWindow: Set sapWindow = wins(1)
  
  'Cleanup edge
  Dim win As stdWindow
  For Each win In stdWindow.CreateManyFromProcessId(EdgeProc.id)
    win.Quit
  Next
  
  
  Call Create.protInit(sapWindow, closeOnDestroy)
End Function
Public Function CreateFromExisting(ByVal win As stdWindow) As sapSAPECC
    Set CreateFromExisting = New sapSAPECC
    Call CreateFromExisting.protInit(win, False)
End Function
Public Function CreateFromAllExisting() As Collection
  Set CreateFromAllExisting = getSapWindows()
End Function

#If OPTION_INCLUDE_ASYNC Then
  'Returns a fiber
  '@returns Object<stdFiber>
  Public Function CreateAsync() As Object
    Dim SAP As sapSAPECC: Set SAP = New sapSAPECC
    Dim fib As stdFiber: Set fib = stdFiber.Create("SAP ECC Client")
    With fib
      Set .Meta("SAP") = SAP
      .add stdCallback.CreateFromObjectMethod(Me, "protFiberStep").bind(1)
      .add stdCallback.CreateFromObjectMethod(Me, "protFiberStep").bind(2)
      .add stdCallback.CreateFromObjectMethod(Me, "protFiberStep").bind(3)
    end with
  End Function
#End If

'@param sapWin - SAP Client window
'@param closeOnDestroy - If true, SAP window will close when the object is destroyed
'@devNote Historic IE-based navigation
'  'When accessing the accessibility model of SAP, for whatever reason when
'  'straying beyond a certain depth, the structure changes meaning that the main sap window swaps from
'  '4.5.4.1 to 4.5.4.2. Thus we first obtain a handle to 4.5.4.1.1
'  Call This.SAPAccMain.CreateFromPath("4.5.4.1.1")
'
'  'Get SAP Main
'  Set zpAccSAPMain = zpAccMain.CreateFromPath("4.5.4.2.4.1.4.1.1.2.1.1.1.1.4.1.4.1.4")
'  If zpAccSAPMain Is Nothing Then Debug.Assert False 'TODO: Fix this
Public Sub protInit(ByVal sapWin As stdWindow, ByVal closeOnDestroy As Boolean)
  If sapWin.Class <> "SAP_FRONTEND_SESSION" Then
    Err.Raise 1, "sapSAPECC::protInit", "Window is not of class SAP_FRONTEND_SESSION"
  End If
  
  'Set SAP Main window
  Set This.SAPMainWindow = sapWin
  Set This.SAPMainAcc = stdAcc.CreateFromHwnd(sapWin.Handle)
  
  'Optimised AccToolbar and AccTransaction bar search:
  With This.SAPMainAcc
    Set This.SAPToolbarAcc = .FindFirst(stdLambda.Create("$1.Name = ""AppToolbar"" and $1.Role = ""ROLE_CLIENT"""))
    Set This.SAPToolbarWindow = stdWindow.CreateFromHwnd(This.SAPToolbarAcc.hwnd)
    
    'Sometimes if you don't do this it bugs out sendkeys etc. 'TOCHECK
    Call This.SAPToolbarWindow.Activate
    
    'Obtain toolbar and transaction bar
    Set This.SAPTransactionAcc = .FindFirst(stdLambda.Create("$1.Role = ""ROLE_COMBOBOX"""))
    Dim errorWin As stdWindow: Set errorWin = This.SAPMainWindow.FindFirst(stdLambda.Create("$1.Class like ""Afx:*:8:00010003:00000010:00000000"""))
    Set This.SAPErrBarAcc = stdAcc.CreateFromHwnd(errorWin.Handle)
  End With
  
  'Wait for infoWindow if it appears
  Call Sleep(100)
  Dim infoWindow As stdWindow: Set infoWindow = sapWindow("Information")
  If Not infoWindow Is Nothing Then
    Dim infoAcc As stdAcc: Set infoAcc = stdAcc.CreateFromHwnd(infoWindow.Handle)
    infoAcc.FindFirst(stdLambda.Create("$1.Name = ""Continue"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
    'Err.Raise 1, "sapSAPECC::protInit", "If you are an admin, this code is unimplemented... If you are a user, wait 15 minutes and try again."
  End If
  
  This.isTemporaryWindow = closeOnDestroy
End Sub

#If OPTION_INCLUDE_ASYNC Then
  'Asynchronous function for fibers
  '@protected
  Public Function protFiberStep(ByVal state as Long, ByVal oFib As Object) As Boolean
    Dim fib As stdFiber: Set fib = oFib
    With fib
      select case state
        case 1
          'Get Existing windows
          Set .Meta("wins") = getSapWindows()
          
          'Create new window via link
          Set .Meta("EdgeProc") = stdProcess.Create("""" & Edge_Exe & """ """ & ECC_URL & """")
          
          'Step to next
          protFiberStep = True
        case 2
          protFiberStep = getSapWindows(.Meta("wins")).count > 0
        case 3
          Dim sapWindow As stdWindow: Set sapWindow = getSapWindows(fib.Meta("wins"))(1)
          Call .Meta("EdgeProc").forceQuit  'Cleanup edge browser
          Call .Meta("SAP").protInit(sapWindow)
          protFiberStep = True
      end select
    End With
  End Function
#End If




Public Property Get Main() As stdAcc
  Set Main = This.SAPMainAcc
End Property

Public Property Get Caption() As String
  Caption = This.SAPMainWindow.Caption
End Property


'Get/Set the current transaction
Public Property Get Transaction() As String
  Transaction = This.SAPTransactionAcc.value
End Property
Public Property Let Transaction(ByVal sTransactId As String)
  This.SAPTransactionAcc.value = sTransactId
  While This.SAPTransactionAcc.value = sTransactId
    'Focus transaction field
    Call This.SAPMainWindow.SendKeysEvent("^/")
    
    'Send enter to execute
    Call This.SAPMainWindow.SendKeysEvent("{Enter}")
    Call Wait(100)
    DoEvents
  Wend
End Property

'Obtain the field in which transactions are get/set
Public Property Get TransactionField() As stdAcc
    Set TransactionField = This.SAPTransactionAcc
End Property

'Obtain the error text reported by SAP
Public Property Get ErrorText() As String
    ErrorText = This.SAPErrBarAcc.name
End Property

'Await a toolbar with a name, if timeout reached `Nothing` is returned
'@param sName - Name of toolbar to wait for
'@returns - Toolbar button if found, else nothing
'@remark - Suggested to use AwaitToolbarName instead
Public Function GetToolbarByName(ByVal sName As String) As stdAcc
  Dim ButtonWin As stdAcc
    For Each ButtonWin In This.SAPToolbarAcc.children
      If ButtonWin.name = sName Then
          Set GetToolbarByName = ButtonWin.children(4)
          Exit Function
      End If
  Next
End Function

'Await a toolbar with a name, if timeout reached `Nothing` is returned
'@param {String} Name of toolbar to wait for
'@returns {stdAcc} Toolbar awaited
Public Function AwaitToolbarName(ByVal sName As String) As stdAcc
  'Wait till name is visible (on the toolbar)
  While True
    'Check all buttons for the name sName, and return if found
    Dim Button As stdAcc: Set Button = GetToolbarByName(sName)
    If Not Button Is Nothing Then
      Set AwaitToolbarName = Button
      Exit Function
    End If
    
    DoEvents
  Wend
End Function

'Await a toolbar with a name, if timeout reached `Nothing` is returned
'@param {String} Name of toolbar to wait for
'@param {Long} Time to wait before a timeout is triggered
'@returns {stdAcc} Toolbar awaited or `Nothing` if timeout is reached.
Public Function AwaitToolbarNameOrError(ByVal sName As String, Optional ByVal iTimeout As Long = 1000, Optional ByVal errorTimeout As Long = 250) As stdAcc
  Dim iStartTime As Long: iStartTime = GetTickCount()
  
  'Wait till name is visible (on the toolbar)
  While True
    'Check all buttons for the name sName, and return if found
    Dim Button As stdAcc: Set Button = GetToolbarByName(sName)
    If Not Button Is Nothing Then Exit Function
    
    'Initiate error check if after half a second since start
    Dim iDiff As Long: iDiff = Abs(GetTickCount() - iStartTime)
    If iDiff >= iErrorTimeout Then
      If Me.ErrorText <> "" Then
        'Check for error information, if found then return nothing
        Set AwaitToolbarNameOrError = Nothing
        Exit Function
      End If
    End If
    
    
    If iDiff >= iTimeout Then
      Exit Function
    End If
    DoEvents
  Wend
End Function


'Send the key combination to execute (note this key combination sometimes changes so you can change it easily)
'@param {String} Key to press
Public Sub cmdExecute(Optional ByVal sKey As String = "F8")
  Call This.SAPMainWindow.SendKeysEvent("{" & sKey & "}")
End Sub

'Send the key combination to go back
Public Sub cmdBack()
  Call This.SAPMainWindow.SendKeysEvent("{F3}")
End Sub

'Send the key combination to exit the current task
Public Sub cmdExit()
  Call This.SAPMainWindow.SendKeysEvent("+{F3}")
End Sub

'Send the key combination to cancel the current task
Public Sub cmdCancel()
  Call This.SAPMainWindow.SendKeysEvent("{F12}")
End Sub

Public Property Get isTimedOut() As Boolean
  Dim pid As Long: pid = This.SAPMainWindow.processID
  Dim wnd As stdWindow, myWnd As stdWindow
  For Each wnd In stdWindow.CreateFromDesktop.children
    If wnd.Exists Then
      If wnd.processID = pid And wnd.Caption <> "" And wnd.Visible Then
        If wnd.Caption Like "*Timeout*" Then
          isTimedOut = True
          Exit Property
        End If
      End If
    End If
  Next
End Property

'Send the key combination to return to the home menu
Public Sub cmdHome()
  While True
    '2024-05-21 Initially we had a search for a window named "SAP LOGON 750" - not sure why...
    
    'Press escape til ErrorText.
    Call This.SAPMainWindow.SendKeysEvent("{Escape}")
    If Me.ErrorText = "Function code cannot be selected" Then Exit Sub
    
    DoEvents
  Wend
  
  Call zpIEWnd.SendKeysEvent("{Escape}", 100)
End Sub

'Send the tab key a number of times
'@param {Long} number of times to send Tab key
Public Sub cmdTab(Optional ByVal iTimes As Long = 1)
  Call This.SAPMainWindow.SendKeysEvent("{TAB " & iTimes & "}")
End Sub

'Focus the SAP toolbar
Public Sub FocusToolbar()
  This.SAPToolbarWindow.Activate
  Dim Button As stdAcc
  Set Button = This.SAPToolbarAcc.children(1)
  Button.Focus = True
End Sub

'Attempt to paste data into the selected control
'@param {String} Text to paste into the field
'@param {String} Key pattern to select all data in the cell. ^a is default, but this could be replaced with e.g. `{Home}+{End}` if ^a isn't supported
Public Sub CopyPasteFieldValue(ByVal sText As String, Optional ByVal sSelectAllKeys As String = "{LCONTROL DOWN}a{LCONTROL UP}", Optional ByVal withConfirmation As Boolean = True)
  stdClipboard.Text = ""
  
  While stdClipboard.Text <> sText
    'Set clipboard to value
    stdClipboard.Text = sText
    
    'Paste value in field
    Call This.SAPMainWindow.SendKeysEvent(sSelectAllKeys, 50)
    Call Sleep(50)
    Call This.SAPMainWindow.SendKeysEvent("{LCONTROL DOWN}v{LCONTROL UP}", 50)
    Call Sleep(50)
    
    'And confirm that it has been pasted
    If withConfirmation Then
      Call This.SAPMainWindow.SendKeysEvent(sSelectAllKeys, 50)
      Call Sleep(50)
      stdClipboard.Text = ""
      Call This.SAPMainWindow.SendKeysEvent("{LCONTROL DOWN}c{LCONTROL UP}", 50)
      Call Sleep(50)
    End If
  Wend
End Sub

'Attempt to copy the data in the selected control
'@param sSelectAllKeys - Key pattern to select all data in the cell. ^a is default, but this could be replaced with e.g. `{Home}+{End}` if ^a isn't supported
'@param tryTimes - The number of attempts before the operation is aborted. When there is a high likelihood this line will be "" it might be worth using a small number here.
'@returns {String} The text copied
Public Function CopySelected(Optional ByVal sSelectAllKeys As String = "{LCONTROL DOWN}a{LCONTROL UP}", Optional ByVal tryTimes As Long = 10) As String
    'Set clipboard to blank
    stdClipboard.Text = ""
    
    'Copy text
    Call This.SAPMainWindow.SendKeysEvent(sSelectAllKeys)
    
    Dim iAttempts As Long: iAttempts = 0
    'Wait for clipboard to be populated with data
    Do While stdClipboard.Text = "" And iAttempts < tryTimes
        iAttempts = iAttempts + 1
        'Copy text
        Call This.SAPMainWindow.SendKeysEvent(sSelectAllKeys)
        Call This.SAPMainWindow.SendKeysEvent("{LCONTROL DOWN}c{LCONTROL UP}", 50)
        DoEvents
        If stdClipboard.Text <> "" Then Exit Do
        Call Sleep(100)
    Loop
    
    'Return data
    CopySelected = stdClipboard.Text
End Function

'Send keys to the main IE window
'@param {String} Keys to send
'@param {Boolean} Whether to send keys in Rawmode. While in rawmode { keys are sent as they are instead of being interpreted as having a special meaning
'@param {Long} Keydelay while sending keys
Public Sub SendKeysMain(ByVal sKeys As String, Optional ByVal rawmode As Boolean = False, Optional ByVal keyDelay As Long = 200)
  Call This.SAPMainWindow.SendKeysEvent(sKeys, rawmode, keyDelay)
End Sub

'Check for a SAP window, if one doesn't exist return nothing.
'@param {String} Name of the window to check for
'@returns {stdWindow} Window object of the new window. Nothing if SAP Window doesn't exist.
Public Property Get sapWindow(ByVal sNameLike As String) As stdWindow
    If This.SAPChildWindow Is Nothing Then
        'TODO: FIX ME
        Dim pid As Long: pid = This.SAPMainWindow.processID
        Dim wnd As stdWindow
        For Each wnd In stdWindow.CreateFromDesktop.children
          If wnd.Exists Then
            If wnd.processID = pid And wnd.Caption Like sNameLike Then
              Set This.SAPChildWindow = wnd
              Exit For
            End If
          End If
        Next
    End If
    
    'Ensure caption like NameLike
    If This.SAPChildWindow Is Nothing Then Exit Property
    If This.SAPChildWindow.Caption Like sNameLike Then
      Set sapWindow = This.SAPChildWindow
    Else
      Set sapWindow = Nothing
    End If
End Property

'Await for a SAP window with a given name
'@param {String} Name of the window to wait for
'@param {Long} Timeout to wait before returning nothing
'@returns {stdWindow} Window object of the new window. If timeout reached `Nothing` is returned
Public Function AwaitSAPWindow(ByVal sName As String, Optional ByVal timeout As Long = 1000, Optional ByVal errorTimeout As Long = 250) As stdWindow
  Dim iStartTime As Long: iStartTime = GetTickCount()
  
  While True
    Dim wnd As stdWindow
    Set wnd = sapWindow("*" & sName & "*")
    If Not wnd Is Nothing Then
      If wnd.Visible Then
        Set AwaitSAPWindow = wnd
        Exit Function
      End If
    End If
    
    'Initiate error check if after half a second since start
    Dim iDiff As Long: iDiff = Abs(GetTickCount() - iStartTime)
    If iDiff >= errorTimeout Then
      'Check for error information, if found then return nothing
      If Me.ErrorText <> "" Then
        Set AwaitSAPWindow = Nothing
        Exit Function
      End If
    End If
    
    If iDiff >= timeout Then
      Set AwaitSAPWindow = Nothing
      Exit Function
    End If
    
    DoEvents
  Wend
  Set AwaitSAPWindow = wnd
End Function

'Quit IE
Public Sub Quit()
  This.SAPMainWindow.Quit
End Sub

'Waits a number of milliseconds
'@param {Long} Number of milliseconds to wait
'@remark Usese Kernel32's Sleep() function
Public Sub Wait(ByVal iMilliseconds As Long)
  Call Sleep(iMilliseconds)
End Sub


'Find and return the workbook who's name is RIIFLO20 (or any pattern)
'@param {String} Find workbooks matching this query pattern
'@returns {Workbook} The workbook identified
Public Function awaitSAPWorkbook(Optional ByVal sQuery As String = "*RIIFLO20*") As Workbook
    'Get sap workbook into private wb
    Set wb = Nothing
    While wb Is Nothing
        Set wb = pGetSAPWorkbook(sQuery)
        DoEvents
    Wend
    Set awaitSAPWorkbook = wb
End Function

'Find and return the workbook who's name matches a pattern
'@param {String} Find workbooks matching this query pattern
'@returns {Workbook} The workbook identified
Private Function pGetSAPWorkbook(ByVal sQuery As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name Like sQuery Then
            Set pGetSAPWorkbook = wb
            Exit Function
        End If
    Next
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  If This.isTemporaryWindow Then Me.Quit
End Sub

'Obtains a collection of existing SAP Client Windows
'@param - Collection of windows to exclude
'@returns - Collection of windows matching SAP Session Client criteria
'@example ```
'Dim col as Collection: set col = getSapWindows()
'set edgeProc = stdProcess.Create(...)
'set col = getSapWindows(col)   'Gets windows excluding those already existing
'```
Private Function getSapWindows(Optional ByVal filterCol As Collection = Nothing) As Collection
  'Creating a filter from param
  Dim filterDict As Object: Set filterDict = CreateObject("Scripting.Dictionary")
  If Not filterCol Is Nothing Then
    Dim filterWin As stdWindow
    For Each filterWin In filterCol
      filterDict(CStr(filterWin.Handle)) = True
    Next
  End If
  
  'Create and filter existing windows
  Dim wins As Collection: Set wins = New Collection
  
  Dim win As stdWindow
  For Each win In stdWindow.CreateFromDesktop().children
    On Error Resume Next
    If win.Exists Then
      If win.Class = "SAP_FRONTEND_SESSION" Then
        If Not filterDict.Exists(CStr(win.Handle)) Then
          wins.add win
        End If
      End If
    End If
  Next
  
  Set getSapWindows = wins
End Function
