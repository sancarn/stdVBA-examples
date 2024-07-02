VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tvAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum tvAccIdentityType
  tvAccIdentityFromEither
  tvAccIdentityFromAccessible
  tvAccIdentityFromPath
End Enum
Public Enum tvAccCreateType
  tvAccCreateByTree
  tvAccCreateByOther
End Enum

Private Type TThis
    acc As stdAcc

    identityType As tvAccIdentityType
    createType As tvAccCreateType
    Parent As tvAcc
    index As Long
End Type
Private This As TThis

'****************
'* Constructors *
'****************

'Creates an `tvAcc` object from an `X` and `Y` point location on the screen.
'@constructor
'@param x - X Coordinate
'@param y - Y Coordinate
'@returns tvAcc - Object at the specified point
Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As tvAcc
  Set CreateFromPoint = Create(stdAcc.CreateFromPoint(x, y))
End Function

'Creates an `tvAcc` object from a window handle.
'@constructor
'@param hwnd - Window handle to wrap
'@returns tvAcc - tvAcc object for the window
Public Function CreateFromHwnd(ByVal hwnd As LongPtr) As tvAcc
  Set CreateFromHwnd = Create(stdAcc.CreateFromHwnd(hwnd))
End Function

'Creates an `tvAcc` object from the current running application (e.g. Excel / Word / Powerpoint).
'@constructor
'@param oApp - Application to create from. Defaults to `Application` if unspecified.
'@returns - IAccessible wrapper around application window.
Public Function CreateFromApplication(Optional ByVal oApp As Object = Nothing) As tvAcc
  Set CreateFromApplication = Create(stdAcc.CreateFromApplication(oApp))
End Function

'Creates an `tvAcc` object from the desktop.
'@constructor
'@returns - IAccessible wrapper around desktop window.
Public Function CreateFromDesktop() As tvAcc
  Set CreateFromDesktop = Create(stdAcc.CreateFromDesktop(), Nothing, 0, tvAccIdentityFromAccessible, tvAccCreateByTree)
End Function

'Creates an `tvAcc` object from an object which implements `IAccessible`.
'@constructor
'@param obj - Object implementing `IAccessible` interface.
'@returns - IAccessible wrapper around an IAccessible object.
Public Function CreateFromIAccessible(ByVal obj As IAccessible) As tvAcc
  Set CreateFromIAccessible = Create(stdAcc.CreateFromIAccessible(obj))
End Function

'Creates an `tvAcc` object for the element the mouse currently hovers over.
'@constructor
'@returns - IAccessible wrapper around element under mouse.
Public Function CreateFromMouse() As tvAcc
  Set CreateFromMouse = Create(stdAcc.CreateFromMouse())
End Function

'Creates an `tvAcc` object for the element at a given path from the current element.
'@constructor
'@param sPath - Path to element, e.g. "P.1.2.3" or "1.2.3". If starting with "P." or ".", these are ignored as they are taken as the current element.
'@returns - IAccessible wrapper from a supplied path.
'@example `Debug.Print tvAcc.CreateFromApplication().CreateFromPath("3.1").name`
Public Function CreateFromPath(ByVal sPath As String) As tvAcc
  'TODO: Implement
End Function

Public Function Create(ByVal acc As stdAcc, Optional ByVal Parent As tvAcc = Nothing, Optional index As Long = 0, Optional ByVal identityType As tvAccIdentityType = tvAccIdentityFromAccessible, Optional ByVal createType As tvAccCreateType = tvAccCreateByOther) As tvAcc
  Set Create = New tvAcc
  Call Create.protInit(acc, Parent, index, identityType, createType)
End Function

'Initialise the object
'@param acc - The accessibility object.
'@protected
Public Sub protInit(ByVal acc As stdAcc, ByVal Parent As tvAcc, ByVal index As Long, ByVal identityType As tvAccIdentityType, ByVal createType As tvAccCreateType)
  Set This.acc = acc
  Set This.Parent = Parent
  This.index = index
  This.identityType = identityType
  This.createType = createType
End Sub

'Return the parent of the IAccessible object
'@returns - Parent of this element
Public Property Get Parent() As tvAcc
  If This.Parent Is Nothing Then
    If This.acc.Parent Is Nothing Then
      Set Parent = Nothing
    Else
      Set Parent = Create(This.acc.Parent, Nothing, 0, tvAccIdentityFromAccessible, tvAccCreateByOther)
    End If
  Else
    Set Parent = This.Parent
  End If
End Property

'Return the children of the IAccessible object
'@returns Collection<stdAcc> - Children of this element
Public Property Get children() As Collection
  Set children = New Collection
  Dim tc As Collection: Set tc = This.acc.children
  Dim acc As stdAcc, index As Long: index = 0
  Dim childIdentityType As tvAccIdentityType
  If This.identityType = tvAccIdentityFromPath Then
    childIdentityType = tvAccIdentityFromPath
  Else
    childIdentityType = tvAccIdentityFromEither
  End If
  
  For Each acc In tc
    index = index + 1
    Call children.Add(Create(acc, Me, index, childIdentityType, tvAccCreateByTree))
  Next
  
End Property

'Get a unique identifier for an element.
'@returns string - Hex string identifier for element.
'@example `stdAcc.CreateFromMouse().Identity`
Public Property Get Identity() As String
  Select Case This.identityType
    Case tvAccIdentityFromEither
      Identity = This.acc.Identity
      If Identity = "Unknown" Then
        This.identityType = tvAccIdentityFromPath
        Identity = This.Parent.Identity & "." & This.index
      Else
        This.identityType = tvAccIdentityFromAccessible
      End If
    Case tvAccIdentityFromAccessible
      Identity = This.acc.Identity
    Case tvAccIdentityFromPath
      Identity = This.Parent.Identity & "." & This.index
  End Select
End Property

'Get the type of identity
'@returns - type of identity
Public Property Get identityType() As tvAccIdentityType
  identityType = This.identityType
End Property

'Return the hwnd of the IAccessible object
'@returns - hwnd of this element
#If VBA7 Then
  Public Property Get hwnd() As LongPtr
#Else
  Public Property Get hwnd() As Long
#End If
    hwnd = This.acc.hwnd
End Property

'Get the path to this element
'@returns - path to this element
Public Function GetPath() As String
  If Not This.Parent Is Nothing Then
    GetPath = This.Parent.GetPath() & "." & This.index
  Else
    If This.createType = tvAccCreateByTree Then
      GetPath = "D"
    Else
      GetPath = "[Unknown]"
    End If
  End If
End Function



'Proxy to stdAcc
Public Property Get HitTest(ByVal x As Long, ByVal y As Long) As stdAcc: HitTest = Create(This.acc.HitTest(x, y)): End Property
Public Property Get Location() As Collection: Set Location = This.acc.Location: End Property
Public Property Get value() As String: value = This.acc.value: End Property
Public Property Let value(val As String): This.acc.value = val: End Property
Public Property Get name() As String: name = This.acc.name: End Property
Public Property Get DefaultAction() As String: DefaultAction = This.acc.DefaultAction: End Property
Public Property Get Role() As String:  Role = This.acc.Role: End Property
Public Property Get States() As String: States = This.acc.States: End Property
Public Property Get StateData() As Long: StateData = This.acc.StateData: End Property
Public Property Get Description() As String: Description = This.acc.Description: End Property
Public Property Get KeyboardShortcut() As String: KeyboardShortcut = This.acc.KeyboardShortcut: End Property
Public Property Get Focus() As Boolean: Focus = This.acc.Focus: End Property
Public Property Get Help() As String: Help = This.acc.Help: End Property
Public Property Get HelpTopic(Optional ByVal sFile As String = "") As String: HelpTopic = This.acc.HelpTopic(sFile): End Property
Public Property Get Text() As String: Text = This.acc.Text: End Property
Public Sub DoDefaultAction(): Call This.acc.DoDefaultAction: End Sub




