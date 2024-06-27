VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Requires:
'* [stdVBA](http://github.com/sancarn/stdVBA)
'    * stdAcc
'    * stdEnumerator
'    * stdLambda
'    * stdProcess
'    * stdWindow
'    * stdICallable

Public Enum EEdgeAwaitStatus
    Successful = 0
    Errored = 1
    TimeoutReached = 2
End Enum

Private Const stdEdgeParamDefaults = "--force-renderer-accessibility --new-window"
Private Type TThis
    Window As stdWindow
    accMain As stdAcc
    accAddressBar As stdAcc
    isTemporaryWindow As Boolean
End Type
Private This As TThis

'******************************************
'*          CONSTRUCTOR METHODS           *
'******************************************

'Create a new Edge window
'@param userParams - User parameters to pass to Edge
'@param closeOnDestroy - If true, the window will close when the object is destroyed, else it will remain open
'@returns - Edge automation object
Public Function Create(Optional ByVal userParams As String = "", optional byval closeOnDestroy as boolean = true) As stdEdge
    'Create application GUID
    Dim sGUID As String: sGUID = getGUID()
    
    'Build command for launch
    Dim sCommand As String: sCommand = """$path"" ""$guid"" $defaultParams $userParams"
    sCommand = Replace(sCommand, "$path", getEdgePath())
    sCommand = Replace(sCommand, "$guid", sGUID)
    sCommand = Replace(sCommand, "$defaultParams", stdEdgeParamDefaults)
    sCommand = Replace(sCommand, "$userParams", userParams)
    
    'Launch process / create new window
    Dim proc As stdProcess: Set proc = stdProcess.Create(sCommand, vbNormalFocus)
    
    'Await window creation
    Dim wnd As stdWindow
    Do
        Set wnd = stdEnumerator.CreateFromIEnumVariant(stdWindow.CreateFromDesktop().children).FindFirst(stdLambda.Create("If $2.Exists Then $2.Caption like $1 & ""*"" else false").Bind(LCase(sGUID)), Nothing)
        DoEvents
    Loop While wnd Is Nothing
    
    'Initialise Edge automation object
    Set Create = New stdEdge
    Call Create.protInit(wnd, closeOnDestroy)
End Function

'Create from an existing window object
'@param wnd - Window to create Edge automation object from
'@param closeOnDestroy - If true, the window will close when the object is destroyed, else it will remain open
'@returns - Edge automation object for the specified window
'@remark Won't be able to force accessibility, therefore will need accessibility assurance otherwise only a finite set of options will be available
Public Function CreateFromExisting(ByVal wnd As stdWindow, Optional ByVal closeOnDestroy as boolean = false) as stdEdge
    Set CreateFromExisting = New stdEdge
    Call CreateFromExisting.protInit(wnd, closeOnDestroy)
End Function

'*************************************
'*          PUBLIC METHODS           *
'*************************************

'Access the Main Window as a stdWindow Object
Public Property Get winMain() As stdWindow
    Set winMain = This.Window
End Property

'Access the Main Window as a stdAcc Object
Public Property Get accMain() As stdAcc
    Set accMain = This.accMain
End Property

'Get/Let Address
Public Property Get Address() As String
    Address = This.accAddressBar.value
End Property
Public Property Let Address(v As String)
    Dim sOldCaption As String: sOldCaption = This.Window.Caption
    This.accAddressBar.value = v
    While InStr(1, This.accAddressBar.States, "STATE_FOCUSED") = 0
        This.Window.Activate
        Call This.Window.SendKeysEvent("{F6}")
        DoEvents
    Wend
    While This.Window.Caption = sOldCaption
        This.Window.Activate
        Call This.Window.SendKeysEvent("{ENTER}")
        DoEvents
    Wend
End Property

'Get/Let isTemporary
'@returns - If set to true, the window will close when the object is destroyed, else it will remain open
Public Property Get isTemporary() As Boolean
    isTemporary = This.isTemporaryWindow
End Property
Public Property Let isTemporary(v As Boolean)
    This.isTemporaryWindow = v
End Property

'Helper calling Address=
'@param sUrl - The url to navigate to
Public Sub Navigate(ByVal sURL As String)
    Address = sURL
End Sub

'Await for a particular condition to be met
'@param awaitFor as stdICallable<stdEdge=>Boolean> - If this callback returns true, the function will immediately set the result of the function to `Successful(=0)`, otherwise will continue awaiting.
'@param error as stdICallable<stdEdge=>Boolean> - If this callback returns true, the function will immediately set the result of the function to `Errored(=1)`, oterwise will continue awaiting.
'@param timeout - Timeout in seconds. SpecialValue: -1 = NoTimeout. When timeout is reached the return value of the function is `TimeoudReached(=2)`
'@returns {EEdgeAwaitStatus} Success status of condition
Public Function AwaitForCondition(awaitFor As stdICallable, Optional error As stdICallable = Nothing, Optional timeout As Long = -1) As EEdgeAwaitStatus
    Dim bCond As Boolean: bCond = True
    Dim bRes As EEdgeAwaitStatus: bRes = EEdgeAwaitStatus.TimeoutReached
    Dim oldTime As Date: oldTime = Now()
    While bCond
        Dim x As stdLambda
        Set x = awaitFor
        If awaitFor.Run(meref()) Then
            bRes = Successful
            bCond = False
        ElseIf timeout <> -1 And Second(Now() - oldTime) > timeout Then
            bRes = TimeoutReached
            bCond = False
        ElseIf Not error Is Nothing Then
            If error.Run(Me) Then
                bRes = Errored
                bCond = False
            End If
        End If
        DoEvents
    Wend
    AwaitForCondition = bRes
End Function

'Await for an acc Element
'@param awaitFor as stdICallable<stdAcc=>Boolean> - If this callback returns true, the function will immediately set the result of the function to `Successful(=0)`, otherwise will continue awaiting.
'@param error as stdICallable<(stdEdge,stdAcc)=>Boolean> - If this callback returns true, the function will immediately set the result of the function to `Errored(=1)`, oterwise will continue awaiting. If none is provided no error is checked.
'@param timeout - Timeout in seconds. Defaults to -1 = NoTimeout. When timeout is reached the return value of the function is `TimeoudReached(=2)`
'@returns - Element awaited for
'@remark calls `Edge.accMain.FindFirst(awaitFor)` to wait for element chosen
Public Function AwaitForAccElement(awaitFor As stdICallable, Optional error As stdICallable = Nothing, Optional timeout As Long = -1) As stdAcc
    If error Is Nothing Then
        Set AwaitForAccElement = This.accMain.AwaitForElement(awaitFor, Nothing, timeout)
    Else
        Set AwaitForAccElement = This.accMain.AwaitForElement(awaitFor, error.Bind(meref()), timeout)
    End If
End Function

'Quit Edge window
Public Sub Quit()
    Call This.Window.Quit
End Sub


'****************************************
'*          PROTECTED METHODS           *
'****************************************

'Initialise object with window
'@param wnd - Window to initialise Edge automation object from
'@param isTemporary - Whether the window should be terminated when the instance is destroyed.
Friend Sub protInit(wnd As stdWindow, isTemporary As Boolean)
    With This
        Set .Window = wnd
        Set .accMain = stdAcc.CreateFromHwnd(wnd.Handle)
        Set .accAddressBar = .accMain.FindFirst(stdLambda.Create("$1.Name like ""Address and search bar"" and $1.Role = ""ROLE_TEXT"""))
        .isTemporaryWindow = isTemporary
    End With
End Sub

'*************************************
'*          EVENT HANDLERS           *
'*************************************

Private Sub Class_Terminate()
    If This.Window.Exists and This.isTemporaryWindow Then Call This.Window.Quit
End Sub

'**************************************
'*          PRIVATE METHODS           *
'**************************************

'Obtain a reference to this object
Private Function meref() As Object
    Set meref = Me
End Function

'Get the file path to Edge
Private Function getEdgePath()
    Const EdgeRegistry = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\"
    getEdgePath = CreateObject("WScript.Shell").RegRead(EdgeRegistry)
End Function


'Obtain a new GUID
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function
