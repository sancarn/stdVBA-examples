VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Enum EUIElementMessage
  uiElementEventChange
  uiElementEventMove
  uiElementEventResize
  uiElementEventMouseClick
  uiElementEventMouseFocus
  uiElementEventMouseMove
  uiElementEventTabSelected
  uiElementEventAddChild
End Enum
Public Enum EUIElementType
  'Group controls
  uiUserForm
  uiFrame
  uiMultiPage

  'Display controls
  uiLabel
  uiImage
  uiTabStrip

  'Edit Controls
  uiTextBox
  uiComboBox
  uiListBox
  uiCheckBox
  uiOptionButton
  uiToggleButton
  uiScrollBar
  uiSpinButton

  'Executors
  uiCommandButton
End Enum

Private pParent  As uiIMessagable
Private pControl As MSForms.Control

Private WithEvents pUserForm       As MSForms.UserForm
Attribute pUserForm.VB_VarHelpID = -1
Private WithEvents pFrame          As MSForms.Frame
Attribute pFrame.VB_VarHelpID = -1
Private WithEvents pMultiPage      As MSForms.MultiPage
Attribute pMultiPage.VB_VarHelpID = -1
Private WithEvents pLabel          As MSForms.Label
Attribute pLabel.VB_VarHelpID = -1
Private WithEvents pImage          As MSForms.Image
Attribute pImage.VB_VarHelpID = -1
Private WithEvents pTabStrip       As MSForms.TabStrip
Attribute pTabStrip.VB_VarHelpID = -1
Private WithEvents pTextBox        As MSForms.TextBox
Attribute pTextBox.VB_VarHelpID = -1
Private WithEvents pComboBox       As MSForms.ComboBox
Attribute pComboBox.VB_VarHelpID = -1
Private WithEvents pListBox        As MSForms.ListBox
Attribute pListBox.VB_VarHelpID = -1
Private WithEvents pCheckBox       As MSForms.CheckBox
Attribute pCheckBox.VB_VarHelpID = -1
Private WithEvents pOptionButton   As MSForms.OptionButton
Attribute pOptionButton.VB_VarHelpID = -1
Private WithEvents pToggleButton   As MSForms.ToggleButton
Attribute pToggleButton.VB_VarHelpID = -1
Private WithEvents pScrollBar      As MSForms.ScrollBar
Attribute pScrollBar.VB_VarHelpID = -1
Private WithEvents pSpinButton     As MSForms.SpinButton
Attribute pSpinButton.VB_VarHelpID = -1
Private WithEvents pCommandButton  As MSForms.CommandButton
Attribute pCommandButton.VB_VarHelpID = -1

Private pType As EUIElementType
Private pid As String

Private pLeft As Double
Private pWidth As Double
'Creating a input component from
Public Function CreateFromControl(ByVal parent As uiIMessagable, ByVal ctrl As Control) As uiElement
  Set CreateFromControl = New uiElement
  Call CreateFromControl.protInit(parent, ctrl)
End Function

Public Function CreateFromType(ByVal parent As uiIMessagable, ByVal oControls As MSForms.Controls, ByVal iCtrlType As EUIElementType, Optional ByVal sName As String, Optional ByVal sCaption As String, Optional ByVal fLeft As Double = 0, Optional ByVal fTop As Double = 0, Optional ByVal fWidth As Double, Optional ByVal fHeight As Double) As uiElement
  Set CreateFromType = New uiElement
  Dim ctrl As Control
  Select Case iCtrlType
    Case uiUserForm:      Set ctrl = oControls.add("Forms.Form.1", sName)
    Case uiFrame:         Set ctrl = oControls.add("Forms.Frame.1", sName)
    Case uiMultiPage:     Set ctrl = oControls.add("Forms.MultiPage.1", sName)
    Case uiLabel:         Set ctrl = oControls.add("Forms.Label.1", sName)
    Case uiImage:         Set ctrl = oControls.add("Forms.Image.1", sName)
    Case uiTabStrip:      Set ctrl = oControls.add("Forms.TabStrip.1", sName)
    Case uiTextBox:       Set ctrl = oControls.add("Forms.TextBox.1", sName)
    Case uiComboBox:      Set ctrl = oControls.add("Forms.ComboBox.1", sName)
    Case uiListBox:       Set ctrl = oControls.add("Forms.ListBox.1", sName)
    Case uiCheckBox:      Set ctrl = oControls.add("Forms.CheckBox.1", sName)
    Case uiOptionButton:  Set ctrl = oControls.add("Forms.OptionButton.1", sName)
    Case uiToggleButton:  Set ctrl = oControls.add("Forms.ToggleButton.1", sName)
    Case uiScrollBar:     Set ctrl = oControls.add("Forms.ScrollBar.1", sName)
    Case uiSpinButton:    Set ctrl = oControls.add("Forms.SpinButton.1", sName)
    Case uiCommandButton: Set ctrl = oControls.add("Forms.CommandButton.1", sName)
  End Select
  If sName <> "" Then ctrl.name = sName
  If sCaption <> "" Then ctrl.Caption = sCaption
  Call CreateFromType.protInit(parent, ctrl)
  CreateFromType.left = fLeft
  CreateFromType.top = fTop
  If fWidth <> 0 Then CreateFromType.width = fWidth
  If fHeight <> 0 Then CreateFromType.height = fHeight
End Function

'Obtain type of input control
Public Property Get uiType() As EUIElementType
  uiType = pType
End Property

'Obtain control object
Public Property Get Control() As Control
  Set Control = pControl
End Property

'Obtain ui object raw item
Public Property Get uiObject() As Object
  Select Case pType
    Case uiUserForm:       Set uiObject = pUserForm
    Case uiFrame:          Set uiObject = pFrame
    Case uiMultiPage:      Set uiObject = pMultiPage
    Case uiLabel:          Set uiObject = pLabel
    Case uiImage:          Set uiObject = pImage
    Case uiTabStrip:       Set uiObject = pTabStrip
    Case uiTextBox:        Set uiObject = pTextBox
    Case uiComboBox:       Set uiObject = pComboBox
    Case uiListBox:        Set uiObject = pListBox
    Case uiCheckBox:       Set uiObject = pCheckBox
    Case uiOptionButton:   Set uiObject = pOptionButton
    Case uiToggleButton:   Set uiObject = pToggleButton
    Case uiScrollBar:      Set uiObject = pScrollBar
    Case uiSpinButton:     Set uiObject = pSpinButton
    Case uiCommandButton:  Set uiObject = pCommandButton
  End Select
End Property

'Get/Set Value
Public Property Get value() As Variant
  Select Case pType
    Case uiTextBox, uiComboBox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollBar, uiSpinButton
      value = uiObject.value
    Case uiLabel
      value = uiObject.Caption
  End Select
End Property
Public Property Let value(v As Variant)
  Select Case pType
    Case uiTextBox, uiComboBox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollBar, uiSpinButton
      uiObject.value = v
    Case uiLabel
      uiObject.Caption = v
  End Select
End Property

'Get/Set Position Left
Public Property Get left() As Double
  left = pLeft
End Property
Public Property Let left(fLeft As Double)
  pLeft = fLeft
  Select Case pType
    Case uiTextBox, uiComboBox, uiListBox, uiOptionButton, uiToggleButton, uiScrollBar, uiSpinButton, uiLabel, uiImage
      uiObject.left = pLeft
    Case uiCheckBox
      uiObject.left = pLeft + IIf(pWidth <= 12, 0, pWidth / 2 - 6)
  End Select
  Call pParent.Message(Me, uiElementEventMove, Array(pLeft, pControl.top))
End Property

'Get/Set Position Top
Public Property Get top() As Double
  top = pControl.top
End Property
Public Property Let top(fTop As Double)
  pControl.top = fTop
  Call pParent.Message(Me, uiElementEventMove, Array(pControl.left, pControl.top))
End Property

'Get/Set Size Width
Public Property Get width() As Double
  width = pWidth
End Property
Public Property Let width(fWidth As Double)
  pWidth = fWidth
  left = pLeft
  uiObject.width = fWidth
  Call pParent.Message(Me, uiElementEventResize, Array(pControl.width, pControl.height))
End Property

'Get/Set Size Height
Public Property Get height() As Double
  height = pControl.height
End Property
Public Property Let height(fHeight As Double)
  pControl.height = fHeight
  Call pParent.Message(Me, uiElementEventResize, Array(pControl.width, pControl.height))
End Property

'Getter for id
Public Property Get id() As String
  id = pid
End Property


'*********************
'* PROTECTED METHODS *
'*********************

'Initialise input control
'@protected
Friend Sub protInit(parent As uiIMessagable, ctrl As Control)
  Set pParent = parent
  Set pControl = ctrl
  pid = getGUID()
  If TypeOf ctrl Is MSForms.UserForm Then
    Set pUserForm = ctrl
    pType = EUIElementType.uiUserForm
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Frame Then
    Set pFrame = ctrl
    pType = EUIElementType.uiFrame
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.MultiPage Then
    Set pMultiPage = ctrl
    pType = EUIElementType.uiMultiPage
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Label Then
    Set pLabel = ctrl
    pType = EUIElementType.uiLabel
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Image Then
    Set pImage = ctrl
    pType = EUIElementType.uiImage
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.TabStrip Then
    Set pTabStrip = ctrl
    pType = EUIElementType.uiTabStrip
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.TextBox Then
    Set pTextBox = ctrl
    pType = EUIElementType.uiTextBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ComboBox Then
    Set pComboBox = ctrl
    pType = EUIElementType.uiComboBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ListBox Then
    Set pListBox = ctrl
    pType = EUIElementType.uiListBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.CheckBox Then
    Set pCheckBox = ctrl
    pType = EUIElementType.uiCheckBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.OptionButton Then
    Set pOptionButton = ctrl
    pType = EUIElementType.uiOptionButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ToggleButton Then
    Set pToggleButton = ctrl
    pType = EUIElementType.uiToggleButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ScrollBar Then
    Set pScrollBar = ctrl
    pType = EUIElementType.uiScrollBar
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.SpinButton Then
    Set pSpinButton = ctrl
    pType = EUIElementType.uiSpinButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.CommandButton Then
    Set pCommandButton = ctrl
    pType = EUIElementType.uiCommandButton
    Exit Sub
  End If
End Sub

Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function






'TODO: Consider calculating and passing X and Y coords of mouse press event
'@remark The following controls are missing the click event: pScrollBar, pSpinButton
Private Sub pUserForm_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pFrame_Click():         Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pLabel_Click():         Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pComboBox_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pListBox_Click():       Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pCheckBox_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pOptionButton_Click():  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pToggleButton_Click():  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pCommandButton_Click(): Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick): End Sub
Private Sub pMultiPage_Click(ByVal index As Long)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventTabSelected, Array(index))
End Sub
Private Sub pTabStrip_Click(ByVal index As Long)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventTabSelected, Array(index))
End Sub
Private Sub pImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick, Array(x, y))
End Sub
Private Sub pTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick, Array(x, y))
End Sub

'Parent object events
Private Sub pUserForm_AddControl(ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, Array(Control, 0)): End Sub
Private Sub pFrame_AddControl(ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, Array(Control, 0)): End Sub
Private Sub pMultiPage_AddControl(ByVal index As Long, ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, Array(Control, index)): End Sub


'Change events
Private Sub pMultiPage_Change():    Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pTabStrip_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pCheckBox_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pComboBox_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pListBox_Change():      Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pOptionButton_Change(): Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pScrollBar_Change():    Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pSpinButton_Change():   Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pTextBox_Change():      Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub
Private Sub pToggleButton_Change(): Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, Array(uiObject.value)): End Sub

'MouseMove events
