VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "biChangeServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Enum IMAGEHLP_STATUS_REASON
    BindOutOfMemory = 0
    BindRvaToVaFailed = 1
    BindVaToRvaFailed = 2
    BindImportByNameFailed = 3
    BindImportByOrdinalFailed = 4
    BindImportImageFailed = 5
    BindImportImageDeferred = 6
    BindExpandFileHeaders = 7
    BindImageComplete = 8
    BindMofComplete = 9
    BindSymbolsNotMatching = 10
    BindReportDebugInfo = 11
    BindReportProgress = 12
    BindReserved = 13
    BindMsgGeneral = 14
    BindMsgCode = 15
    BindMsgValue = 16
    BindMsgItem = 17
    BindMsgPhase = 18
    BindMsgProgress = 19
    BindMsgDebug = 20
    BindMsgNTSTATUS = 21
    BindMsgMof = 22
    BindMaxStatus = 23
End Enum

#If Win64 then
  Const NULL_PTR as LongLong = 0^
#Else
  Const NULL_PTR as Long = 0&
#End if

#If VBA7 Then
  Private Declare PtrSafe Function UnDecorateSymbolName Lib "imagehlp.dll" (ByVal DecoratedName As LongPtr, ByVal UnDecoratedName As LongPtr, UnDecoratedLength As Long, ByVal Flags As Long) As Long
  Private Declare PtrSafe Function ImageHlpStatusRoutine Lib "imagehlp.dll" (ByVal Reason As IMAGEHLP_STATUS_REASON, ByVal ImageName As LongPtr, ByVal DllName As LongPtr, ByVal Va As LongPtr, ByVal ParameterSize As Long, ByVal Parameter As LongPtr ) As Long
  Private Declare PtrSafe Function MapAndLoad Lib "imagehlp.dll" ( ByVal ImageName As LongPtr, ByVal DllPath As LongPtr, ByRef LoadedImage As Any, ByVal Attr As Long, ByVal CallBack As LongPtr) As Long
  Private Declare PtrSafe Function UnMapAndLoad Lib "imagehlp.dll" (ByRef LoadedImage As Any) As Long ' Returns True (Non-zero) on success, False (0) on failure
  Private Declare PtrSafe Function SymInitialize Lib "imagehlp.dll" ( ByVal hProcess As LongPtr, ByVal UserSearchPath As LongPtr, ByVal fInvadeProcess As Long ) As Long
  Private Declare PtrSafe Function SymCleanup Lib "imagehlp.dll" ( ByVal hProcess As LongPtr ) As Long
  Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32.dll" () As LongPtr
  'TODO: Is this required? Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
#Else
  Enum LongPtr
    [_]
  End Enum
  Private Declare Function UnDecorateSymbolName Lib "imagehlp.dll" (ByVal DecoratedName As LongPtr, ByVal UnDecoratedName As LongPtr, UnDecoratedLength As Long, ByVal Flags As Long) As Long
  Private Declare Function ImageHlpStatusRoutine Lib "imagehlp.dll" (ByVal Reason As IMAGEHLP_STATUS_REASON, ByVal ImageName As LongPtr, ByVal DllName As LongPtr, ByVal Va As LongPtr, ByVal ParameterSize As Long, ByVal Parameter As LongPtr ) As Long
  Private Declare Function MapAndLoad Lib "imagehlp.dll" ( ByVal ImageName As LongPtr, ByVal DllPath As LongPtr, ByRef LoadedImage As Any, ByVal Attr As Long, ByVal CallBack As LongPtr) As Long
  Private Declare Function UnMapAndLoad Lib "imagehlp.dll" (ByRef LoadedImage As Any) As Long ' Returns True (Non-zero) on success, False (0) on failure
  Private Declare Function SymInitialize Lib "imagehlp.dll" ( ByVal hProcess As LongPtr, ByVal UserSearchPath As LongPtr, ByVal fInvadeProcess As Long ) As Long
  Private Declare Function SymCleanup Lib "imagehlp.dll" ( ByVal hProcess As LongPtr ) As Long
  Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As LongPtr
  'TODO: Is this required? Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
#End If

Private Type LOADED_IMAGE
  ModuleName as String      'PSTR   'TODO: Test, Might have to be LongPtr
  hFile as LongPtr          'HANDLE
  MappedAddress as LongPtr  'Ptr to UCHAR
  FileHeader as LongPtr     'Ptr to IMAGE_NT_HEADERS
  rest(56) as byte          'Total Size either 88 or less. Doesn't matter really, just needs to be big enough. We only need first 4 pointers for this algorithm
End Type

Private Type IMAGE_EXPORT_DIRECTORY
  Characteristics as Long
  TimeDateStamp as Long
  MajorVersion as Integer
  MinorVersion as Integer
  Name as Long
  Base as Long
  NumberOfFunctions as Long
  NumberOfNames as Long
  AddressOfFunctions as Long
  AddressOfNames as Long
  AddressOfNameOrdinals as Long
End Type

Private Type IMAGE_DATA_DIRECTORY
  VirtualAddress as Long
  Size as Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
  Magic As Integer
  MajorLinkerVersion as Byte
  MinorLinkerVersion as Byte
  SizeOfCode as Long
  SizeOfInitializedData as Long
  SizeOfUninitialisedData as Long
  AddressOfEntryPoint as Long
  BaseOfCode as Long
  #If Win64 then
    ImageBase as LongLong
  #Else
    BaseOfData as Long
    ImageBase as Long
  #End If
  SectionAlignment as Long
  FileAlignment as Long
  MajorOperatingSystemVersion as Integer
  MinorOperatingSystemVersion as Integer
  MajorImageVersion as Integer
  MinorImageVersion as Integer
  MajorSubsystemVersion as Integer
  MinorSubsystemVersion as Integer
  Win32VersionValue as Long
  SizeOfImage as Long
  SizeOfHeaders as Long
  CheckSum as Long
  SubSystem as Integer
  DllCharacteristics as Integer
  SizeOfStackReserve as LongPtr
  SizeOfStackCommit as LongPtr
  SizeOfHeapReserve as LongPtr
  SizeOfHeapCommit as LongPtr
  LoaderFlags as Long
  NumberOfRvaAndSizes as Long
  DataDirectory() as IMAGE_DATA_DIRECTORY
End type

Private Type IMAGE_FILE_HEADER
  Machine As Integer
  NumberOfSections As Integer
  TimeDateStamp As Long
  PointerToSymbolTable As Long
  NumberOfSymbols As Long
  SizeOfOptionalHeader As Integer
  Characteristics As Integer
End Type

Private Type IMAGE_NT_HEADERS
  Signature as Long
  FileHeader as IMAGE_FILE_HEADER
  OptionalHeader as IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_NT_HEADERS64 
    Signature As Long
    ' IMAGE_FILE_HEADER (portion)
    
    ' IMAGE_OPTIONAL_HEADER64 (portion - only what's commonly needed for MapAndLoad's output)
    Magic As Integer 
    AddressOfEntryPoint As Long
    ImageBase As LongPtr
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As LongPtr
    SizeOfStackCommit As LongPtr
    SizeOfHeapReserve As LongPtr
    SizeOfHeapCommit As LongPtr
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    ' DataDirectory(0 to IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) would go here
    ' For simplicity, we often just access ImageBase and SizeOfImage after MapAndLoad
End Type



Const IMAGE_FILE_MACHINE_I386 as Long = &H014c
Const IMAGE_FILE_MACHINE_AMD64 as Long = &H8664


Type TImageFunction
  Name as string
  EntryPoint as LongPtr
  Ordinal as LongPtr
  DemangledName as string
End Type
Type TImageData
  Initialised as boolean
  ModuleName as string
  TotalFunctions as Long
  NamedFunctions as Long
  OrdBase as Long
  Bitness as Long
  Functions() as TImageFunction
End Type
Private Type TThis
    path as string
    image as TImageData
End Type
Private This as TThis

Public Function Create(ByVal path as string) as stdDLL
    set Create = new stdDLL
    Call Create.protInit(path)
End Function

Public Sub protInit(ByRef path as string)
    This.path = path
End Sub

'Obtain module name from module data
Public Property Get Name() as String
  If not This.image.Initialised then Call InitialiseModuleData()
  Name = This.image.ModuleName
End Property

Public Property Get Bitness() as Integer
  If not This.image.Initialised then Call InitialiseModuleData()
  Name = This.image.Bitness
End Property

'Get the function definitions for this DLL
'@returns Collection<Dictionary<string, variant>> - Collection of dictionaries containing function data.
Public Property Get Functions() as Collection
  If not This.image.Initialised then Call InitialiseModuleData()
  static cache as Collection
  if cache is nothing then
    set cache = new Collection

    Dim i as Long
    With This.image
      For i = lbound(.Functions) to ubound(.Functions)
        With .Functions(i)
          Cache.Add CreateDict("Name", .Name, "EntryPoint", .EntryPoint, "Ordinal", .Ordinal, "DemangledName", .DemangledName)
        End With
      next
    end with
  End if
  set Functions = Cache
End Property



Private Sub InitialiseModuleData()
  'TODO: Is this required? 
  'Dim hImageHlp as LongPtr: hImageHlp = LoadLibrary("imagehlp.dll")

  Dim bPath() as Byte: bPath = StrConv(this.path & Chr(0), vbFromUnicode)
  Dim image as LOADED_IMAGE
  if MapAndLoad(VarPtr(bPath(0)), NULL_PTR, image, 1, 1) <> 0 then
    This.image.ModuleName = image.ModuleName
    Dim Machine as Integer: Machine = NumGet(image.FileHeader + 4, "ushort")
    if Machine = IMAGE_FILE_MACHINE_I386 or Machine = IMAGE_FILE_MACHINE_AMD64 Then
      'MappedAddress    ==> image.MappedAddress
      'IMAGE_NT_HEADERS ==> image.FileHeader
      Dim size as LongPtr
      Dim ExportDirPtr as LongPtr: ExportDirPtr = ImageDirectoryEntryToData(image.MappedAddress, 0, 0, size)
      if ExportDirPtr <> 0 then
        Dim ExportDir as IMAGE_EXPORT_DIRECTORY: 
        'TBC...
      end if
    end if
    
    
    This.image.initialised = true
    Call UnMapAndLoad(image)
  end if
End Sub



'Undecorates/Demangles a VC++ DLL Function Name
'@param sDecorated - The decorated function name
'@example ```
'Debug.Print UndecorateFunction("??0OsfAuthenticationApiHelper@@QEAA@AEBV?$weak_ptr@VOfficeExtensionManager@@@std@@@Z")
''=> public: __cdecl OsfAuthenticationApiHelper::OsfAuthenticationApiHelper(class std::weak_ptr<class OfficeExtensionManager> const & __ptr64) __ptr64
'```
Private Function UndecorateFunction(ByVal sDecorated As String) As String
  Const BUFFER_MAX As Long = 2048
  Dim bDecorated() As Byte: bDecorated = StrConv(sDecorated & Chr(0), vbFromUnicode)
  Dim bUndecorated() As Byte: ReDim bUndecorated(0 To BUFFER_MAX)
  
  Const UNDNAME_COMPLETE As Long = 0&
  Dim iLength As Long: iLength = UnDecorateSymbolName(VarPtr(bDecorated(0)), VarPtr(bUndecorated(0)), BUFFER_MAX, UNDNAME_COMPLETE)
  UndecorateFunction = Left(StrConv(bUndecorated, vbUnicode), iLength)
End Function
