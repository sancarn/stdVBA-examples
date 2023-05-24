Attribute VB_Name = "xrExtractor"
Public Sub ExtractorMain()
  Dim ePaths As stdEnumerator: Set ePaths = stdEnumerator.CreateFromListObject(dataPaths.ListObjects("Paths"))
  Dim categories As xrCategories: Set categories = xrCategories.Create(dataCategories.ListObjects("Categories"))
  Dim rules As xrRules: Set rules = xrRules.Create(dataRules.ListObjects("Rules"))
  
  Dim results As stdArray: Set results = stdArray.Create()
  
  Dim oPath
  For Each oPath In ePaths
    Dim wb As Workbook: Set wb = Workbooks.Open(oPath("Path"))
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
      'Get category from sheet data
      Dim sCategory As String: sCategory = categories.getCategory(ws)
      
      'Perform extraction
      If sCategory <> "" Then Call results.Push(rules.executeRules(ws, sCategory))
    Next
    
    wb.Close
  Next
  
  Call exportResults(dataOutput, "Output", results)
End Sub



'@param {Worksheet} Sheet to export data to
'@param {stdEnumerator<Dictionary>} Results to export to range
Sub exportResults(ByVal ws As Worksheet, ByVal sTableName As String, results As stdArray)
  Dim vFields: vFields = results.item(1).keys()
  Dim iFieldLength As Long: iFieldLength = UBound(vFields) - LBound(vFields) + 1
  
  Dim vResults() As Variant
  ReDim vResults(1 To results.Length + 1, 1 To iFieldLength)
  Dim iResCol As Long: iResCol = 0
  Dim vField
  For Each vField In vFields
    iResCol = iResCol + 1
    vResults(1, iResCol) = vField
  Next
  
  Dim iRow As Long
  For iRow = 1 To results.Length
    iResCol = 0
    For Each vField In vFields
      iResCol = iResCol + 1
      vResults(iRow + 1, iResCol) = results.item(iRow)(vField)
    Next
  Next
  
  dataOutput.UsedRange.Clear
  With dataOutput.Range("A1").Resize(results.Length + 1, iFieldLength)
    .value = vResults
    .WrapText = False
    With dataOutput.ListObjects.add(xlSrcRange, .Cells)
      .name = sTableName
    End With
  End With
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
