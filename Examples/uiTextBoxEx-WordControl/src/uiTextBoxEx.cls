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

Public Function Create(fr As MSForms.Frame) As uiTextBoxEx
  Set Create = New uiTextBoxEx
  Call Create.protInit(fr)
End Function

Public Property Get Text() As String
  Text = pAppWordDoc.Content.Text
End Property
Public Property Let Text(RHS As String)
  pAppWordDoc.Content.Text = RHS
End Property

Friend Sub protInit(fr As MSForms.Frame)
  'Set word control
  Set pWndUserformFrame = stdWindow.CreateFromIUnknown(fr)
  
  'Set bindings to word
  Set pAppWord = CreateObject("Word.Application")
  pAppWord.Visible = True
  Set pAppWordDoc = pAppWord.documents.Add()
  
  With pAppWord.ActiveWindow
    'Customise window properties
    .View.Type = 6 'webLayout
    
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
  pProcWord.forceQuit
End Sub
