Attribute VB_Name = "xrExtractor"
Public Sub ExtractorMain()
  Application.StatusBar = "Setting up..."
  
  Dim ePaths As stdEnumerator: Set ePaths = stdEnumerator.CreateFromListObject(dataPaths.ListObjects("Paths"))
  Dim categories As xrCategories: Set categories = xrCategories.Create(dataCategories.ListObjects("Categories"))
  Dim rules As xrRules: Set rules = xrRules.Create(dataRules.ListObjects("Rules"))
  
  Dim results As stdArray: Set results = stdArray.Create()
  
  'Boot a new instance of the excel application
  'opening workbooks in a seperate instance ensures that the workbook isn't loaded on each loop cycle.
  'it will also ensure that we can set various performance increasing settings, allowing us to scan the
  'workbook as soon as possible. Finally we can prevent annoyances like Asking to update links.
  'note we intentionally don't set this app to visible.
  Dim xlApp As Excel.Application: Set xlApp = New Excel.Application
  xlApp.AskToUpdateLinks = False
  xlApp.ScreenUpdating = False
  xlApp.AutomationSecurity = msoAutomationSecurityForceDisable
  xlApp.EnableEvents = False
  xlApp.DisplayAlerts = False
  xlApp.Calculation = xlCalculationManual
  
  'Loop over each path in the supplied paths table
  Dim oPath, i As Long: i = 0
  For Each oPath In ePaths
    'Increment progress counter
    i = i + 1
    
    'Only extract if Processed <> Yes
    If oPath("Processed") <> "Yes" Then
      'Open workbook in hidden instance
      Dim wb As Workbook: Set wb = xlApp.Workbooks.Open(oPath("Path"))
      
      Dim ws As Worksheet
      For Each ws In wb.Worksheets
        Application.StatusBar = "Processing workbook " & i & "/" & ePaths.Length & " Worksheet: " & ws.name
        
        'Get category from sheet data
        Dim sCategory As String: sCategory = categories.getCategory(ws)
        
        'Perform extraction
        If sCategory <> "" Then Call results.Push(rules.executeRules(ws, sCategory))
      Next
      
      wb.Close SaveChanges:=False
      
      'Update row
      Call setRowCell(oPath, "Processed", "Yes")
    End If
  Next
  
  'Close hidden app instance
  xlApp.Quit
  
  
  Application.StatusBar = "Exporting results..."
  Call exportResults(dataOutput, "Output", results)
  Application.StatusBar = Empty
End Sub


'@param {Worksheet} Sheet to export data to
'@param {(stdEnumerator|stdArray)<Dictionary>} Results to export to range
Sub exportResults(ByVal ws As Worksheet, ByVal sTableName As String, results As Object)
  If TypeOf results Is stdArray Or TypeOf results Is stdEnumerator Then
    If results.Length = 0 Then Exit Sub 'length guard
    
    'Populate headers in result
    Dim vResults() As Variant
    Dim vFields: vFields = results.item(1).keys()
    Dim iFieldLength As Long: iFieldLength = UBound(vFields) - LBound(vFields) + 1
    ReDim vResults(1 To results.Length + 1, 1 To iFieldLength)
    Dim iResCol As Long: iResCol = 0
    Dim vField
    For Each vField In vFields
      iResCol = iResCol + 1
      vResults(1, iResCol) = vField
    Next
    
    'Populate data in result
    Dim iRow As Long
    For iRow = 1 To results.Length
      iResCol = 0
      For Each vField In vFields
        iResCol = iResCol + 1
        vResults(iRow + 1, iResCol) = results.item(iRow)(vField)
      Next
    Next
    
    'Write data to output
    dataOutput.UsedRange.Clear
    With dataOutput.Range("A1").Resize(results.Length + 1, iFieldLength)
      .value = vResults
      .WrapText = False
      With dataOutput.ListObjects.add(xlSrcRange, .Cells)
        .name = sTableName
      End With
    End With
  End If
End Sub

'Given a row object supplied by `stdEnumerator.CreateFromListObject()`, update a cell's value based on it's column name
'@param {Object} Row object
'@param {String} Column name
'@param {Variant} Value to set cell to
Private Sub setRowCell(ByVal oRow As Object, ByVal sColumnName As String, ByVal vValue As Variant)
  Application.Intersect(oRow("=ListRow").Range, oRow("=ListColumns")(sColumnName).Range).value = vValue
End Sub



'@test
Sub test_exportResults()
  Dim res As stdArray: Set res = stdArray.Create(td(1, 2, 3, 4), td(4, 5, 6, 7))
  Call exportResults(dataOutput, "Output", res)
End Sub

'@test
'@helper
Private Function td(a, b, c, d) As Object
  Dim o As Object
  Set o = CreateObject("Scripting.Dictionary")
  o("a") = a
  o("b") = b
  o("c") = c
  o("d") = d
  Set td = o
End Function
