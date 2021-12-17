VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTextBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Bindings to userform
Private pWndUserformFrame As stdWindow

'Bindings to word application
Private pAppWord As Object
Private pAppWordDoc As Object
Private pWndWord As stdWindow
Private pWndWordPane As stdWindow
Private pProcWord As stdProcess

'Variables to store display settings
Private pSettingsReadOnly As Boolean
Private pSettingsAnnotations As Boolean
Private pSettingsRuler As Boolean

Public Function Create(fr As MSForms.Frame) As uiTextBoxEx
  Set Create = New uiTextBoxEx
  Call Create.protInit(fr)
End Function
Public Sub Terminate()
  Set pWndWordPane.parent = pWndWord
  pAppWord.Quit False
  DoEvents
End Sub

'Obtain the text provided in the word control.
'@returns {String} The text in the word control, note this text is unformatted
Public Property Get Text() As String
  Text = pAppWordDoc.Content.Text
End Property
Public Property Let Text(RHS As String)
  pAppWordDoc.Content.Text = RHS
  If pSettingsReadOnly = True Then
    'Applies Read Only protection (Type = wdAllowOnlyReading = 3)
    'To Unprotect:   pAppWordDoc.Unprotect
    pAppWordDoc.Protect Type:=3
  End If
End Property

'Get/Set the OpenXML data of the word document. This allows users to save and load formatted text to and from the control.
'@returns {String} OpenXML document containing word data
Public Property Get OpenXML() As String
  OpenXML = pAppWordDoc.Content.XML
End Property
Public Property Let OpenXML(RHS As String)
  With pAppWordDoc.Content
    .Delete
    .InsertXML(RHS)
  End With
  If pSettingsReadOnly = True Then
    'Applies Read Only protection (Type = wdAllowOnlyReading = 3)
    'To Unprotect:   pAppWordDoc.Unprotect
    pAppWordDoc.Protect Type:=3
  End If
End Property

Public Property Get ReadOnly() As Boolean
  ReadOnly = pSettingsReadOnly
End Property
Public Property Let ReadOnly(ReadOnlyState As Boolean)
  'Read-Only protection is not set until after the text
  'property has been set
  pSettingsReadOnly = ReadOnlyState
End Property

Public Property Get DisplayRuler() As Boolean
  DisplayRuler = pSettingsRuler
End Property
Public Property Let DisplayRuler(DisplayState As Boolean)
  pSettingsRuler = DisplayState
  pAppWord.ActiveWindow.ActivePane.DisplayRulers = pSettingsRuler
End Property

Public Property Get ShowAnnotations() As Boolean
  ShowAnnotations = pSettingsAnnotations
End Property
Public Property Let ShowAnnotations(ShowAnnotationsState As Boolean)
  pSettingsAnnotations = ShowAnnotationsState
  pAppWord.ActiveWindow.ActivePane.View.ShowAll = pSettingsAnnotations
End Property

'Initialise word document
Friend Sub protInit(fr As MSForms.Frame)
  'Set word control
  Set pWndUserformFrame = stdWindow.CreateFromIUnknown(fr)
  
  'Set bindings to word
  Set pAppWord = CreateObject("Word.Application")
  pAppWord.Visible = False
  Set pAppWordDoc = pAppWord.documents.Add()
  
  With pAppWord.ActiveWindow
    'Customise window properties
    .View.Type = 6 'webLayout
    
    'If the ActiveWindow is in a Maximised state, then the
    'code returns an error for each of the following two lines.
    .WindowState = 0

    'Set width and height to that of the frame
    .Width = pWndUserformFrame.Width - 30
    .Height = pWndUserformFrame.Height
  
    'Bind class to ActiveWindow object
    Set pWndWord = stdWindow.CreateFromHwnd(.hwnd)
  End With
  
  'Get window pane object (there is probably is a better way of finding this)
  Set pWndWordPane = pWndWord.FindFirst(stdLambda.Create("$1.Caption = """""))
  
  'Bind word pane to frame and set position
  With pWndWordPane
    'Set the parent of the word frame to that of the Frame
    Set .parent = pWndUserformFrame
    
    'Change x, y, width and height
    'Note: changing width and height alters X and Y so alter x and y afterwards
    .X = 0
    .Y = 0
    .Width = pWndUserformFrame.Width
    .Height = pWndUserformFrame.Height
    
    
    Debug.Print .X & "," & .Y & "," & .Width & "," & .Height
  End With
  
  'Hide word window
  pWndWord.Visible = False
  
  'Bind to process
  Set pProcWord = stdProcess.CreateFromProcessId(pWndWord.ProcessID)
  
  'Focus window
  pAppWord.ActiveWindow.Activate
End Sub

'Upon termination of the class we need to properly cleanup the word application, as it's still running.
Private Sub Class_Terminate()
  'Force quit process
  If Not pProcWord Is Nothing Then
    pProcWord.forceQuit
  End If
End Sub
