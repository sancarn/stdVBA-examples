VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sapSAPECCIH06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const EXPORT_DIR As String = "C:\Users\$username\Documents\MY SAP OUTPUTS"

'Requires:
'* stdICallable
'* stdAcc
'* stdWindow
'* stwSAPECC
'  * stdICallable
'  * stdAcc
'  * stdWindow
'  * stdClipboard
Public Enum ESAPECCIH06State
    s1_Initialised2
    s2_AwaitingResults
    s3_ObtainedResults
    s4_ObtainingCharacteristicsOptions
    s5_ObtainingCharacteristicsData
    s6_OutFilePopulating
    s7_Cleanup
    s8_Finished
End Enum

Private Type TSingleton
  hasGuiXT As VbTriState
  wasGuiXTEnabled As Boolean
End Type

Private Type SAPECCIH06Internals
    State As ESAPECCIH06State
    VariantName As String
    Characteristics As String
    OutName As String
    OutFile As Workbook
    accSpreadsheetButton As stdAcc
    wndChars As stdWindow
    wndExport As stdWindow
End Type

Private Type TThis
  Singleton As TSingleton
  ecc As stwSAPECC
  Internals As SAPECCIH06Internals
End Type
Private This As TThis

'Create an instance of IH06 synchronously
'@constructor
'@returns {sapSAPECCIH06} An IH06 search object
Public Function CreateSync() As sapSAPECCIH06
  'Save existing status
  This.Singleton.wasGuiXTEnabled = settingGuiXT
  settingGuiXT = True
  
  Dim ecc As stwSAPECC: Set ecc = stwSAPECC.Create()
  Set CreateSync = New sapSAPECCIH06
  Call CreateSync.protInit(ecc)
End Function

'Create an instance of IH06 from an existing
'@constructor
'@param Ecc - An instance of SAP ECC
'@returns {sapSAPECCIH06} An IH06 search object
Public Function CreateFromExisting(ByVal ecc As stwSAPECC) As sapSAPECCIH06
  Set CreateFromExisting = New sapSAPECCIH06
  Call CreateFromExisting.protInit(ecc)
End Function

'Returns true if GuiXT was enabled when creating the object
Public Property Get protWasGuiXTEnabled() As Boolean
  protWasGuiXTEnabled = This.Singleton.wasGuiXTEnabled
End Property


'Initialisation function
'@protected
'@param {stwSAPECC} An instance of SAP ECC which can be used for automation
Friend Sub protInit(ByVal ecc As stwSAPECC)
  Set This.ecc = ecc
  This.ecc.Transaction = "IH06"
End Sub

Public Property Get ecc() As stwSAPECC
    Set ecc = This.ecc
End Property

Public Property Get State() As ESAPECCIH06State
    State = This.Internals.State
End Property

'Load a variant with a specified name
'@param Name - The name of the variant to load
Public Sub loadVariant(ByVal name As String)
  If Not This.ecc.Caption Like "*Display Functional Location: Functional Location Selection*" Then Err.Raise 1, "stdSAPECCIH06::loadVariant", "Incorrect screen for loading variants"
  Call btnPressGetVariant
  Dim variantWnd As stdWindow: Set variantWnd = This.ecc.AwaitSAPWindow("Find Variant", 2000)
  With variantWnd
    .Activate
    Call This.ecc.Wait(200)
    Call CopyPasteFieldValue(variantWnd, name)
    Call This.ecc.Wait(100)
    Call .SendKeysEvent("{DOWN}{DOWN}")
    Call This.ecc.Wait(100)
    Call CopyPasteFieldValue(variantWnd, "")
    Call This.ecc.Wait(100)
    
    Dim getVarExecButt As stdAcc: Set getVarExecButt = stdAcc.CreateFromHwnd(variantWnd.handle).FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Execute"""))
    If Not This.ecc.RetryWaitTimeout( _
      stdLambda.Create("$1.DoDefaultAction").Bind(getVarExecButt), _
      stdLambda.Create("not $1.Visible").Bind(variantWnd) _
      , 500) Then
      '... panic ...?
      Stop
    End If
    
  End With
End Sub



'Await for the results window
Public Sub awaitResults()
  Call This.ecc.AwaitToolbarName("Spreadsheet")
End Sub

'Load a classification
'@param sName - The name of the classification / characteristic class to show
Public Sub loadClassification(ByVal sName As String)
  Dim i As Long
  For i = 1 To 4
    Call menuPressShowHideClassification
    
    Dim wndChars As stdWindow: Set wndChars = This.ecc.AwaitSAPWindow("Class Limitation * Entries")
    
    Dim iTry As Long
    For iTry = 1 To 3
      If wndChars.Visible Then Exit For
      DoEvents
      Call This.ecc.Wait(100)
    Wend
    
    If wndChars.Visible Then Exit For
  Next
  
  'Click the "Restrict Values" button in the window to select a specific name
  Dim accCharsWnd As stdAcc: Set accCharsWnd = stdAcc.CreateFromHwnd(.wndChars.handle)
  Call accCharsWnd.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Restrict Values""")).DoDefaultAction
  
  'Wait for window to change
  While accCharsWnd.name Like "Class Limitation * Entr*": DoEvents: Wend
  Call This.ecc.Wait(100)
  
  'Restrict values to the name specified
  Call CopyPasteFieldValue(accCharsWnd, sName)
  Call accCharsWnd.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.name = ""Continue""")).DoDefaultAction
  
  'Await window to change back to "Class Limitation etc."
  While Not accCharsWnd.name Like "Class Limitation * Entr*": DoEvents: Wend
  
  'Triggering select all, will select only the filtered values, i.e. this specific class
  Call accCharsWnd.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.name = ""Select all""")).DoDefaultAction
  
  'Trigger the "Copy" button (execute button)
  Call accCharsWnd.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.name = ""Copy""")).DoDefaultAction
End Sub

Public Sub pop()
  Select Case t
      
      Case ESAPECCIH06State.s5_ObtainingCharacteristicsData
          'Check if progress should be made
          Dim bProgress As Boolean
          bProgress = Not .wndChars.Visible
          '
          If bProgress Then
              Call exportAsSpreadsheet2(.OutName)
              .State = s7_Cleanup
          End If
      Case ESAPECCIH06State.s7_Cleanup
          Call This.ecc.cmdHome
          .State = s8_Finished
  End Select
End Sub

Public Sub Wait_Temp(ByVal iMS As Long)
    Call This.ecc.Wait(iMS)
End Sub

Public Property Get isFinished() As Boolean
    isFinished = This.Internals.State = ESAPECCIH06State.s8_Finished
End Property

Public Sub btnPressExecute()
    This.ecc.Transaction = "/8"
End Sub
Public Sub btnPressGetVariant()
    This.ecc.Transaction = "/17"
End Sub
Public Sub menuPressShowHideClassification()
    This.ecc.Transaction = "/Menu=6,14"
    '/Menu=6,14/CAPITAL_PROJ_DATA
End Sub
Public Sub btnPressSelectAllResults()
    This.ecc.Transaction = "/5"
End Sub
Public Sub btnPressCalculator()
  This.ecc.Transaction = "/16"
End Sub
'Public Function exportAsSpreadsheet() As Workbook
'    Call btnPressCalculator
'
'    'Await and handle spreadsheet exporter window
'    Dim wnd As stdWindow: Set wnd = This.ecc.AwaitSAPWindow("Excel: Number of Key*")
'    Dim acc As stdAcc: Set acc = stdAcc.CreateFromHwnd(wnd.handle)
'    Call acc.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Continue""")).DoDefaultAction
'
'    'Await and Handle Table/Pivot table window
'    While Not wnd.Caption Like "Export list object to XXL":  DoEvents: Wend
'    Call wnd.SendKeysEvent("{UP}")
'    Call acc.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Continue""")).DoDefaultAction
'
'    'Save workbooks before spreadsheet export, to find new spreadsheet workbooks generated
'    Dim eOldWBs As stdEnumerator
'    Set eOldWBs = stdEnumerator.CreateFromIEnumVariant(Application.Workbooks).Map(stdLambda.Create("$1.Name"))
'
'    'Press continue to start export to excel
'    While Not wnd.Caption Like "Export list object*":  DoEvents: Wend
'    Call acc.FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Continue""")).DoDefaultAction
'
'    'Save for future use
'    Set This.Internals.wndExport = wnd
'
'    'Find workbook not within existing workbooks
'    Dim sNewWB As String
'    While sNewWB = ""
'        sNewWB = stdEnumerator.CreateFromIEnumVariant(Application.Workbooks) _
'            .Map(stdLambda.Create("$1.Name")) _
'            .FindFirst(stdLambda.Create("not $1.includes($2)").Bind(eOldWBs), Empty)
'        DoEvents
'    Wend
'    Set exportAsSpreadsheet = Application.Workbooks(sNewWB)
'End Function

'Export spreadsheet as file name
'@param sFileName - The file name to download the file as
Public Function exportAsSpreadsheet(ByVal sFileName As String)
  'List > Save > File...
  This.ecc.Transaction = "/Menu=1,11,3"
  Dim wnd As stdWindow: Set wnd = This.ecc.AwaitSAPWindow("Save list in file*")
  With wnd
    Call .SendKeysEvent("{DOWN}")
    With stdAcc.CreateFromHwnd(.handle)
      .FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Continue""")).DoDefaultAction
      
      '
      Do
          DoEvents
      Loop Until wnd.Visible And wnd.Caption Like "Display Functional Location:*"
      
      Dim sGUID As String: sGUID = getGUID()
      Dim sFileLocation As String: sFileLocation = Replace(EXPORT_DIR, "$username", Environ("username")) & sGUID & ".tsv"
      Call CopyPasteFieldValue(wnd, sGUID & ".tsv")
      .FindFirst(stdLambda.Create("$1.Role = ""ROLE_PUSHBUTTON"" and $1.Name = ""Generate""")).DoDefaultAction
    End With
    
    While .Visible: DoEvents: Wend
    Do Until This.ecc.ErrorText Like "* bytes transmitted*"
      DoEvents
    Loop
    
    Dim sOutFile As String: sOutFile = ThisWorkbook.path & "\" & sFileName & ".tsv"
    If DIR(sOutFile) <> "" Then Kill sOutFile
    
    'Rename file
    Name sFileLocation As ThisWorkbook.path & "\" & sFileName & ".tsv"
  End With
End Function

Private Sub Cancel()
    This.ecc.Transaction = "/12"
End Sub

'Create a randomly generated GUID
'@returns - A guid
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function

'Get/Let If GuiXT is enabled
'@returns - True if GuiXT was enabled, false otherwise.
Public Property Get settingGuiXT() As Boolean
  If Not hasGuiXT Then Exit Function
  settingGuiXT = stdReg.CreateFromItem("HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize\GuiXT").value = 1
End Property
Public Property Let settingGuiXT(ByVal v As Boolean)
  Call stdReg.CreateFromItem("HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize\GuiXT").SetValue(IIf(v, 1, 0), Value_DWORD)
End Property

'Determine if GuiXT registry is present
Private Property Get hasGuiXT() As Boolean
  If This.Singleton.hasGuiXT = vbUseDefault Then
    Dim reg As stdReg
      For Each reg In stdReg.CreateFromKey("HKEY_CURRENT_USER\Software\SAP\SAPGUI Front\SAP Frontend Server\Customize").children
      If reg.name = "GuiXT" Then
        This.Singleton.hasGuiXT = vbTrue
      End If
    Next
  End If
  hasGuiXT = This.Singleton.hasGuiXT
End Property

Private Sub Class_Terminate()
  'If GuiXT wasn't enabled at start then remove it
  If Not sapSAPECCIH06.protWasGuiXTEnabled Then settingGuiXT = False
End Sub



'Attempt to paste data into the selected control
'@param {String} Text to paste into the field
'@param {String} Key pattern to select all data in the cell. ^a is default, but this could be replaced with e.g. `{Home}+{End}` if ^a isn't supported
Private Sub CopyPasteFieldValue(ByVal wnd As stdWindow, ByVal sText As String, Optional ByVal sSelectAllKeys As String = "{LCONTROL DOWN}a{LCONTROL UP}", Optional ByVal withConfirmation As Boolean = True)
  stdClipboard.Text = ""
  
  While stdClipboard.Text <> sText
    'Set clipboard to value
    stdClipboard.Text = sText
    
    'Paste value in field
    Call wnd.SendKeysEvent(sSelectAllKeys, 50)
    Call This.ecc.Wait(50)
    Call wnd.SendKeysEvent("{LCONTROL DOWN}v{LCONTROL UP}", 50)
    Call This.ecc.Wait(50)
    
    'And confirm that it has been pasted
    If withConfirmation Then
      Call wnd.SendKeysEvent(sSelectAllKeys, 50)
      Call This.ecc.Wait(50)
      stdClipboard.Text = ""
      Call wnd.SendKeysEvent("{LCONTROL DOWN}c{LCONTROL UP}", 50)
      Call This.ecc.Wait(50)
    End If
  Wend
End Sub
