VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcExcelRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Private Function dsISrc_getName() As String
  dsISrc_getName = "ExcelRange"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Set dsISrc_linkFiber = template
  With template
    .Meta("file") = Data.item("File")
    .Meta("sheet") = Data.item("Sheet")
    .Meta("range") = Data.item("Range")
    .Meta("OutFile") = destPath
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessExcel").Bind(1), "1. Finding workbook")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessExcel").Bind(2), "2. Opening workbook")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessExcel").Bind(3), "3. Connecting to workbook")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessExcel").Bind(4), "4. Dumping to CSV")
  End With
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("File")
      .add "DisplayText", "Source File"
      .add "Tooltip", "The location of the Excel file to get the data from."
      .add "Type", "File"
      With .AddObject("File")
        .add "Extensions", "*.xlsm,*.xlsx,*.xlsb"
      End With
    End With
    
    With .AddObject("Sheet")
      .add "DisplayText", "Source Sheet"
      .add "Tooltip", "The sheet from which to get the data from."
      .add "Type", "Dropdown"
      With .AddObject("Dropdown")
        .add "Options", stdCallback.CreateFromObjectMethod(Me, "protGetSheetNamesCollection")
      End With
    End With
    
    With .AddObject("Range")
      .add "DisplayText", "Range address"
      .add "Tooltip", "The range from which to get the data from."
      .add "Type", "Text"
    End With
  End With
End Function



Public Function protGetSheetNamesCollection(ByVal Data As stdJSON) As Collection
  Dim cSheets As Collection: Set cSheets = New Collection
  Dim sFile As String: sFile = Data.item("File")
  If sFile <> "" Then
    Dim wb As Workbook: Set wb = Workbooks.open(sFile)
    Dim ws As Worksheet: For Each ws In wb.Worksheets
      cSheets.add ws.name
    Next
    wb.Close False
  End If
  Set protGetSheetNames = cSheets
End Function

'Processes an excel query.
'@fiberRunner
'@protected
Public Function protProcessExcel(ByVal stage As Long, ByVal fiber As stdFiber) As Boolean
  Dim wb As Workbook, xlApp As Excel.Application
  With fiber
    Select Case stage
      Case 1 'Single stage as currently all processing is synchronous. TODO: Explore asynchronous options.
        'Obtain sanatised workbook name
        Dim wbPath As String: wbPath = fiber.Meta("file")
        wbPath = Replace(wbPath, "\", "/") 'sanatise path
        wbPath = Replace(wbPath, "%20", " ") 'sanatise path
        Dim iName As Long: iName = InStrRev(wbPath, "/")
        Dim wbName As String: wbName = Mid(wbPath, iName + 1)
        
        On Error GoTo 0
        On Error Resume Next
        Set wb = Workbooks(wbName)
        If Not wb Is Nothing Then
          Set .Meta("wb") = wb
          protProcessExcel = True
        Else
          'Workbook not already open, therefore open it in a new process
          Const CmdTemplate As String = "excel.exe /e ""$1"""
          .Meta("Process") = shell(Replace(CmdTemplate, "$1", wbPath), vbMinimizedNoFocus)
          protProcessExcel = True
        End If
      Case 2 ' await windows on process ID if required
        If .Meta.Exists("Process") Then
          protProcessExcel = stdWindow.CreateManyFromProcessId(.Meta("Process")).count > 0
        Else
          protProcessExcel = True
        End If
      Case 3 ' get workbook from process ID
        If .Meta.Exists("Process") Then
          Dim xlWin As stdWindow: Set xlWin = stdWindow.CreateManyFromProcessId(.Meta("Process")).item(1)
          Dim objWin As stdWindow: Set objWin = xlWin.FindFirst(stdLambda.Create("$1.Class = ""EXCEL7"""))
          Set xlApp = objWin.AsIAccessible.Application
          Set .Meta("app") = xlApp
          Set .Meta("wb") = xlApp.Workbooks(1)
        End If
        protProcessExcel = True
      Case 4
        'Get excel app and workbook. Note Application.Intersect doesn't work on foreign ranges
        If .Meta.Exists("Process") Then
          Set xlApp = .Meta("app")
        Else
          Set xlApp = Application
        End If
        Set wb = .Meta("wb")
        
        
        'Obtain sheet
        On Error Resume Next
        Dim shName As String: shName = .Meta("sheet")
        Dim sheet As Worksheet: Set sheet = wb.Sheets(shName)
        If sheet Is Nothing Then
          Call fiber.RaiseCriticalError("Worksheet '" & shName & "' in file " & wbPath & " not found")
          protProcessExcel = True
          Exit Function
        End If
        
        'Obtain range
        Dim sRng As String: sRng = .Meta("range")
        Dim rng As Range: Set rng = xlApp.Intersect(sheet.UsedRange, sheet.Range(sRng))
        If rng Is Nothing Then
          Call fiber.RaiseCriticalError("Range '" & sRng & "' for table '" & .Meta("Table") & "' not found")
          protProcessExcel = True
          Exit Function
        End If
        
        'Dump to CSV
        Call RangeToCSV(.Meta("OutFile"), rng)
        
        'Cleanup
        Call wb.Close(False)
        If .Meta.Exists("Process") Then Call xlApp.Quit
        protProcessExcel = True
    End Select
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function



