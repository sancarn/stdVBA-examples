VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec: 
'A class used to automate  objects which implement IAccessible
'
'
'CONSTRUCTORS
'    [ ] CreateFromApplication(ByVal xlApp as Excel.Application)
'    [ ] CreateFromWindow(ByVal hwnd as LongPtr)
'    [ ] CreateFromPID(ByVal processID as Long)
'
'Properties
'    [ ] Application() as Excel.Application
'    [ ] EditMode                                                                    'via Accessibility
'    [ ] MacroRecordMode                                                             'via Accessibility
'    [ ] VBAMode() as EVBAMode                                                       'aka EbMode

#If VBA7 = 0 then
  Public Enum LongPtr
    [_]
  End Enum
#End If

#If Win64 Then
  Const NULL_PTR as LongLong = 0^
#Else
  Const NULL_PTR As Long = 0&
#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long


Public Enum EXLVBAMode
  Running
  Design
End Enum

Public Enum EXLEditMode
  Undefined 'This value is does not correlate with an actual Excel cell edit mode, but is used as a return value when an exception occurs while attempting to poll the current state.
  Ready     'Represents the default state. 
  Enter     'Indicates when a cell is selected and typing has begun, or when F2 is keyed twice
  Edit      'Indicates in-cell editing, which occurs as a result of double clicking a cell, keying F2 to enter data.
  Point     'Indicates in-cell formula selection mode, which occurs as a result of entering a formula, and clicking the cells/ranges to include in that formula.
End Enum

Private Type TEbMode
  initialised as boolean
  hVBE as LongPtr
  hEbMode as LongPtr
End Type
Private Type TThis
  hwnd as LongPtr
  Application as Object
  EbMode as TEbMode
End Type
Private This as TThis



Public Function CreateFromApplication(ByVal xlApp as Excel.Application) as xlApplication
  set CreateFromApplication = new xlApplication
  Call CreateFromApplication.protInit(xlApp)
End Function

Public Function CreateFromHWND(ByVal hwnd as LongPtr) as xlApplication
  ' Obtain the Excel.Application object from the hwnd
    Dim sIID as string: sIID = "00020400-0000-0000-C000-000000000046"
    Dim IID_IDispatch as GUID
    If IIDFromString(StrPtr(sIID), IID_IDispatch) = S_OK then
      Dim xlApp as object: 
      if AccessibleObjectFromWindow(hwnd, &HOBJID_NATIVEOM, IID_IDispatch, xlApp) = 0 then
        set CreateFromHWND = CreateFromApplication(xlApp)
      else
        Err.Raise 1, "", ""
      End if
    Else
      Err.Raise 1, "", ""
    end if
End Function

Public Sub protInit(ByVal xlApp as Application)
  set This.Application = xlApp
  This.hwnd = xlApp.hwnd
End Sub

Public Property Get Application as Excel.Application
  set Application = This.Application
End Property

'@remark Uses LPenHelper function of XLCall32.dll https://www.reddit.com/r/excel/comments/10d6l98/a_novel_technique_for_monitoring_excels_edit_mode/
'@remark also see https://stackoverflow.com/questions/76067288/monitoring-excels-edit-mode-via-nativewindow-subclassing-not-working-in-excel-d
Public Property Get EditMode as EXLEditMode
  if this.Application is Application then
    'TODO: Use LPenHelper
  else
    'TODO: Use stdAcc "Status Bar" Pane > "" pane > "" pane > "Status Bar" Status Bar > "Cell Mode *" Button
  end if
End Property

'@remark Searches and returns EbMode from VBA module. See https://github.com/thetrik/VbTrickTimer for example implementation
Public Property Get VBAMode as EXLVBAMode
  #If VBA7 And Win64 Then
    Const sVBEModule as string = "VBE7"
  #ElseIf VBA6 Then
    Const sVBEModule as string = "VBE6"
  #Else
    Const sVBEModule as string = "VBA6"
  #End If

  if not This.EbMode.initialised then
    'Step 1: obtain a handle to VBE7 module.
    if xlApp is Application then
      This.EbMode.hVBE = GetModuleHandle(StrPtr(sVBEModule))
    else
      Dim pID as Long: Call GetWindowThreadProcessId(This.hwnd, pID)
      Dim hProcess as LongPtr: hProcess = OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, false, pID)
      Dim hModules(1 to 1024) as LongPtr, iModuleCount as Long
      if(EnumProcessModules(hProcess, VarPtr(hModules(1)), lenb(hModules), iModuleCount) = 1) then
        Dim sModuleCheck as string: sModuleCheck = "*" & sVBEModule & ".dll"
        Dim i As long
        For i = 1 to iModuleCount
          Dim sModulePath as string: sModulePath = String(255," ")
          if GetModuleFileName(hProcess, hModules(i), sModulePath, 255) = 1 then
            if trim(sModulePath) like sModuleCheck then
              This.EbMode.hVBE = hModules(i)
              Exit For
            end if
          end if
        next
      end if
    end if
    if This.EbMode.hVBE = 0 then Err.Raise 1, "", "Cannot find loaded VBE7.dll"
    Dim hVBE as LongPtr: hVBE = This.EbMode.hVBE
    
    'VBA6+ hides/embeds EbMode function into VBE dlls
    #If VBA6 Or VBA7 Then
      'Step 2: Obtain information to help us search through the VBE module in order to find address of EbMode function.  See the following for more information (S:?): https://learn.microsoft.com/en-us/windows/win32/debug/pe-format 
      Dim e_lfanew  As Long:    Call CopyMemory(e_lfanew,  ByVal hVBE + &H3C, 4)            'logical file address for the New Executable (S:Module?) header
      Dim iNumOfSec As Integer: Call CopyMemory(iNumOfSec, ByVal hVBE + e_lfanew + 6   , 2) 'Number of sections in module
      Dim iOptSize  As Integer: Call CopyMemory(iOptSize,  ByVal hVBE + e_lfanew + &H14, 2) 'Size of optional header(S:?)
      Dim pSection as LongPtr: pSection = hVbe + e_lfanew + &H18 + iOptSize                 'Pointer to the section see section headers here: https://learn.microsoft.com/en-us/windows/win32/debug/pe-format#section-table-section-headers
      Dim lIndex as Long
      For lIndex = 0 To iNumOfSec - 1
          Dim cName As Currency: Call CopyMemory(cName, ByVal pSection, 8)
          
          'Search for ".text" section
          If cName = 50023612.1134@ Then
              CopyMemory pStartScan, ByVal pSection + &HC, 4   'Obtain VirtualAddress of executable image to start scan for EbMode at...
              CopyMemory pEndScan, ByVal pSection + &H8, 4     'Obtain VirtualSize of executable image to end scan for EbMode at...
              pStartScan = pStartScan + hVbe
              pEndScan = pEndScan + pStartScan - 1
              Exit For
          End If
          
          pSection = pSection + &H28  'Each section is 40 bytes long so this skips a section
      Next
      If pStartScan = 0 Or pEndScan = 0 Then Err.Raise 1, "", "Cannot find virtual address of section containing EbMode."

      'Step 3: Search for Proc/MethCallEngine thunk template
      #if VBA7 And Win64 then
        'Method call engine template:
        '48 89 4C 24 08 48 89 54 24 10 4C 89 44 24 18 4C
        '89 4C 24 20 48 B8 11 11 11 11 11 11 11 11 48 0B
        'C0 74 32 48 B8 XX XX XX XX XX XX XX XX FF D0 48      <== XX XX XX XX XX XX XX XX - EbMode  @ 37 byte offset (&H25)
        '83 F8 02 74 20 48 B8 11 11 11 11 11 11 11 11 48
        '8B 4C 24 08 48 8B 54 24 10 4C 8B 44 24 18 4C 8B
        '4C 24 20 FF E0 48 33 C0 C2 11 11 YY YY YY YY YY
        Const EbModeOffset as Long = &H25
        Const sTemplateHex = "48 89 4C 24 08 48 89 54 24 10 4C 89 44 24 18 4C 89 4C 24 20 48 B8 11 11 11 11 11 11 11 11 48 0B C0 74 32 48 B8 00 00 00 00 00 00 00 00 FF D0 48 83 F8 02 74 20 48 B8 11 11 11 11 11 11 11 11 48 8B 4C 24 08 48 8B 54 24 10 4C 8B 44 24 18 4C 8B 4C 24 20 FF E0 48 33 C0 C2 11 11" ' 00 00 00 00 00
        Const sMaskHex     = "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF" ' 00 00 00 00 00
      #Else
        'Method call engine template (32 bit):
        'A1 YY YY YY YY 0B C0 74 13 B8 XX XX XX XX FF D0         <== XX XX XX XX - EbMode  @ 10 byte offset (&HA)
        '83 F8 02 74 07 B8 11 11 11 11 FF E0 33 C0 C2 11
        '11
        Const EbModeOffset as Long = &HA
        Const sTemplateHex = "A1 00 00 00 00 0B C0 74 13 B8 00 00 00 00 FF D0 83 F8 02 74 07 B8 11 11 11 11 FF E0 33 C0 C2 11 11"
        Const sMaskHex     = "FF 00 00 00 00 FF FF FF FF FF 00 00 00 00 FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
      #End If

      'Setup template and mask
      Dim bTemplate() As Byte: bTemplate = getBytesFromHex(sTemplateHex)
      Dim bMask() As Byte:     bMask     = getBytesFromHex(sMaskHex)

      'Map array to data
      Dim tSAMap As SAFEARRAY
      tSAMap.cbElements = 1
      tSAMap.cDims = 1
      tSAMap.fFeatures = FADF_AUTO
      tSAMap.Bounds.cElements = CLng(pEndScan - pStartScan) + 1
      tSAMap.pvData = pStartScan

      
      Dim bData() As Byte
      DupArray bData, VarPtr(tSAMap)

      'Find the start of the template within bData array
      lIndex = FindSignature(bData(), bTemplate(), bMask())
      DupArray bData, 0@ 'Free data
      If lIndex = -1 Then Err.Raise 1, "", "Cannot find EbMode signiature in virtual address data."

      'Copy EbMode function address into hEbMode
      Call CopyMemory(This.EbMode.hEbMode, ByVal pStartScan + lIndex + EbModeOffset, Len(This.EbMode.hEbMode))
    #Else
      This.EbMode.hEbMode = GetProcAddress(hVbe, "EbMode")
    #End If
    if hProcess <> 0 Then Call CloseHandle(hProcess)
    This.EbMode.initialised = true
  end if

  VBAMode = iif(CallFuncByPtr(This.EbMode.hEbMode) = 1, Running, DesignMode)
End Property


Private Function FindSignature(ByRef bData() As Byte, ByRef bSignature() As Byte, ByRef bMask() As Byte) As Long
  Dim lDataIndex  As Long: lDataIndex = 0
  Dim lSignIndex  As Long: lSignIndex = 0
  
  Do While lDataIndex <= UBound(bData)
    If bData(lDataIndex) = bSignature(lSignIndex) Or bMask(lSignIndex) = 0 Then
      lSignIndex = lSignIndex + 1
      If lSignIndex > UBound(bSignature) Then
        FindSignature = lDataIndex - UBound(bSignature)
        Exit Function
      End If
    Else
      If lSignIndex Then
        lDataIndex = lDataIndex - lSignIndex + 1
        lSignIndex = 0
      End If
    End If
    lDataIndex = lDataIndex + 1
  Loop
  
  FindSignature = -1
End Function

Private Function getBytesFromHex(ByVal sHex As String) As Byte()
  Dim b() As Byte
  Dim s() As String
  s = Split(sHex, " ")
  ReDim b(0 To UBound(s))
  Dim i As Long
  For i = 0 To UBound(s)
    b(i) = CByte("&H" & s(i))
  Next
  getBytesFromHex = b
End Function