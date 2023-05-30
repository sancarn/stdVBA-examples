VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xrLambdaEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A small regex based DSL compiling to `stdLambda`.
'@param {String} DSL syntax to execute. DSL is superset of stdLambda with addition in remarks.
'@returns {stdLambda} Compiled lambda to execute.
'@remarks Expects global "targetSheet" late-binding. Replaces `r[...]` with `targetSheet.range("...")` and `[...]` with `targetSheet.range("...").value`
Public Function Create(ByVal sLambda As String) As stdLambda
  'Obtain extension library
  Static oFuncExt As Object: If oFuncExt Is Nothing Then Set oFuncExt = GetExtensionLibrary()
  
  'DSL translations
  sLambda = stdRegex.Create("r\[(?<range>\w+)\]").Replace(sLambda, "targetSheet.range(""$range"")")            'r[...] ==> targetSheet.Range("...")
  sLambda = stdRegex.Create("\[(?<range>\w+)\]").Replace(sLambda, "targetSheet.range(""$range"").value")       ' [...] ==> targetSheet.Range("...").value
  
  'Return translated lambda
  Set Create = stdLambda.Create(sLambda)
  Set Create.oFunctExt = oFuncExt
End Function

'Get a dictionary containing all methods of this object as stdCallback
'@returns {Object<Dictionary<string, stdCallback>>} Dictionary of public methods and callbacks pointing to them
Private Function GetExtensionLibrary() As Object
  Static oLib As Object
  If oLib Is Nothing Then
    Set oLib = CreateObject("Scripting.Dictionary")
    
    'Add formula functions
    Dim vMethodName
    For Each vMethodName In stdCOM.Create(Application.WorksheetFunction).Methods
      Set oLib(vMethodName) = stdCallback.CreateFromObjectMethod(Application.WorksheetFunction, vMethodName)
    Next
    
    'Add methods from this class
    For Each vMethodName In stdCOM.Create(Me).Methods
      If vMethodName <> "Create" Then
        Set oLib(vMethodName) = stdCallback.CreateFromObjectMethod(Me, vMethodName)
      End If
    Next
    
    'Add keywords
    Set oLib("stdRegex") = stdRegex
    Set oLib("targetSheet") = Nothing
    oLib("targetCategory") = ""
    
    'Add extensions
    Dim extension
    For Each extension In stdEnumerator.CreateFromListObject(dataExtensions.ListObjects("Extensions")).AsCollection
      Set oLib(extension("Namespace")) = Application.Run("'" & extension("Path") & "'!getExtension")
    Next
    ThisWorkbook.Activate
  End If
  Set GetExtensionLibrary = oLib
End Function

'------------------------------------------------------------------------------------------------------------

'Generates a new GUID
'@returns {string} a new random GUID
Public Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Obtains a collection of shapes which intersect with a range
'@param {Range} The range to search for shapes within
'@param {Optional Boolean=False} Exit on the first shape found. Set this to true if you're only interested in finding 1 shape.
'@returns {Collection<Shape>} A collection of all intersecting shapes
Public Function getIntersectingShapes(ByVal rng As Range, Optional exitOnFind As Boolean = False) As Collection
  Dim shp As Shape
  Dim shp_x1 As Double, shp_y1 As Double, shp_x2 As Double, shp_y2 As Double
  Dim rng_x1 As Double, rng_y1 As Double, rng_x2 As Double, rng_y2 As Double
  Dim bInBound As Boolean, cRet As Collection
  Set cRet = New Collection
  
  For Each shp In rng.Parent.shapes
    'Shp bounds
    shp_x1 = shp.Left
    shp_y1 = shp.Top
    shp_x2 = shp.Left + shp.Width
    shp_y2 = shp.Top + shp.Height
    
    'Rng bounds
    rng_x1 = rng.Left
    rng_y1 = rng.Top
    rng_x2 = rng.Left + rng.Width
    rng_y2 = rng.Top + rng.Height
    
    'Is top left point within bounding box?
    bInBound = ((shp_x1 >= rng_x1 And shp_x1 <= rng_x2) _
               And (shp_y1 >= rng_y1 And shp_y1 <= rng_y2)) _
               Or ((shp_x2 >= rng_x1 And shp_x2 <= rng_x2) _
               And (shp_y2 >= rng_y1 And shp_y2 <= rng_y2))
    
    'If in the bound then image exists and exit function
    If bInBound Then
      cRet.add shp
      If exitOnFind Then
        Set getIntersectingShapes = cRet
        Exit Function
      End If
    End If
  Next
  Set getIntersectingShapes = cRet
End Function

'Exports a single shape to a PNG file.
'@param {Shape} Shape to export to file as PNG
'@param {String} Directory to export file to
'@param {String} New file name
'@returns {String} The full file path of the newly created PNG
Public Function exportShapeToFile(ByVal shp As Shape, ByVal sFileDir As String, ByVal sFileName As String) As String
  Dim sFile As String: sFile = sFileDir & "\" & IIf(right(sFileDir, 1) = "\", "", "\") & sFileName & ".png"
  Call stdPicture.CreateFromShape(shp).saveAs(sFile, stdPicFormatPNG)
  exportShapeToFile = sFile
End Function

'Exports a collection of shapes to a PNG files.
'@param {Collection} Collection of shape objects to export to folder of PNG files
'@param {String} Base directory to create export folder within
'@param {String} Base name of new files
'@returns {String} The path to the folder containing all exported images
'@example ShapeCollectionToFiles(getIntersectingShapes(r[A1:M24]), "C:\Temp", "Image")
Public Function exportShapeCollectionToFiles(ByVal shapes As Collection, ByVal sFileDir As String, ByVal sBaseName As String) As String
  If shapes.Count > 0 Then
    'Create folder to store images
    Dim sDir As String: sDir = sFileDir & IIf(right(sFileDir, 1) = "\", "", "\") & getGUID()
    MkDir sDir
    
    'Export each shape to file
    Dim shp As Shape, s As String, i As Long: i = 0
    For Each shp In shapes
      i = i + 1
      Call stdPicture.CreateFromShape(shp).saveAs(sDir & "\" & sBaseName & "_" & i & ".png", stdPicFormatPNG)
    Next
    exportShapeCollectionToFiles = sDir
  End If
End Function



'Like `WorksheetFunction.Offset` but considers merged cells, which can be very useful when working with structured data formats.
'@param {Range} The range to perform the offset on
'@param {Long=0} Number of merged rows to offset by
'@param {Long=0} Number of merged columns to offset by
'@returns {Range} The resulting range after offsetting by merged rows and columns.
Public Function OffsetMerged(ByVal rng As Range, Optional ByVal rows As Long = 0, Optional ByVal cols As Long = 0) As Range
  Dim sheet As Worksheet: Set sheet = rng.Parent
  Dim isDown As Boolean: isDown = rows > 0
  Dim isRight As Boolean: isRight = cols > 0
  Dim iOffset As Long, i As Long
  
  'If this cell is merged then make the full range the full merge area
  If rng.MergeCells Then Set rng = rng.MergeArea
  
  'Offset rows
  For i = 1 To Abs(rows)
    'Different procedures if going up or down
    If isDown Then
      Set rng = sheet.Cells(rng.rows.Count + rng.Row, rng.Column)
    Else
      Set rng = sheet.Cells(rng.Row - 1, rng.Column)
    End If
    
    'If this cell is merged then make the full range the full merge area
    If rng.MergeCells Then Set rng = rng.MergeArea
  Next
  
  'Offset columns
  For i = 1 To Abs(cols)
    If isRight Then
      Set rng = sheet.Cells(rng.Row, rng.Columns.Count + rng.Column)
    Else
      Set rng = sheet.Cells(rng.Row, rng.Column - 1)
    End If
    
    'If this cell is merged then make the full range the full merge area
    If rng.MergeCells Then Set rng = rng.MergeArea
  Next
  
  'Return value cell of range
  Set OffsetMerged = rng.Cells(1, 1)
End Function







'Test function, adds 2 numbers
Public Function add(a As Double, b As Double) As Double
  add = a + b
End Function

'--Testing--
'@test
Friend Sub test(ByVal sLambda As String)
  With Create(sLambda)
    Set .oFunctExt("targetSheet") = ActiveSheet
    MsgBox .Run()
  End With
End Sub

