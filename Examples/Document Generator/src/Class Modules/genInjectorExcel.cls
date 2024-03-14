VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "genInjectorExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements genIInjector

Private Type TThis
  templatePath As String
  templateWorkbook As Workbook
  templateSheet As Worksheet
  targetWorkbook As Workbook
  targetSheet As Worksheet
End Type
Private This As TThis

'Create a new instance of the injector
'@param templatePath - The path to the template file
'@returns - An instance of a genIInjector object - within here include all initialisation stuff
Public Function Create(ByVal templatePath As String, ByVal templateSheet As String) As genIInjector
  Dim gen As New genInjectorExcel
  Call gen.protInit(templatePath, templateSheet)
  Set Create = gen
End Function

'Initialise the injector
'@param templatePath - The path to the template file
'@param templateSheet - The name of the sheet in the template file
Public Sub protInit(ByVal templatePath As String, ByVal templateSheet As String)
  This.templatePath = templatePath
  Set This.templateWorkbook = Workbooks.open(templatePath)
  Set This.templateSheet = This.templateWorkbook.Sheets(templateSheet)
End Sub


'Obtain the bindings for the injector
'@returns stdEnumerator<Object<Dictionary<"lambda", string, "setter", stdICallable>>> - a list of bindings for the injector
'@example ```
'  For each binding in injector.getFormulaBindings()
'    binding("setter").call(stdLambda.create(binding("lambda")).call(row))
'  Next
'```
Public Function getFormulaBindings() As stdEnumerator
  Dim bindings As Collection: Set bindings = FindAll(This.templateSheet, "*$*")
  Dim tmp As Collection: Set tmp = New Collection
  For Each cell In bindings
    Dim tmpCell As Range: Set tmpCell = cell.resize(1, 1)
    Call tmp.add(CreateDictionary( _
      "lambda", tmpCell.value, _
      "setter", stdCallback.CreateFromObjectMethod(Me, "setValue").Bind(tmpCell.address), _
      "getSetterTarget", stdCallback.CreateFromObjectMethod(Me, "getCell").Bind(tmpCell.address), _
      "meta", CreateDictionary() _
    ))
  Next
  Set getFormulaBindings = stdEnumerator.CreateFromIEnumVariant(tmp)
End Function


'Initialise the target workbook
Public Function InitialiseTarget() As Object
  Set This.targetWorkbook = Workbooks.add
  This.templateSheet.Copy Before:=This.targetWorkbook.Sheets(1)
  Application.DisplayAlerts = False
  This.targetWorkbook.Sheets(2).Delete
  Application.DisplayAlerts = True
  Set This.targetSheet = This.targetWorkbook.Sheets(1)
  Application.ScreenUpdating = False
  
  'Set the target to be the workbook itself
  Set InitialiseTarget = This.targetWorkbook
End Function

'Cleanup the target workbook
Public Sub CleanupTarget()
  Call This.targetWorkbook.Close(False)
  Application.ScreenUpdating = True
End Sub

'Delegate for "setter" property
Public Sub setValue(ByVal address As String, ByVal value As Variant)
  Dim cell As Range: Set cell = This.targetSheet.Range(address)
  If IsObject(value) Then
    Select Case TypeName(value)
      Case "Dictionary"
        If value.Exists("Type") Then
          Select Case value("Type")
            Case "Picture"
              Dim picture As Object: Set picture = This.targetSheet.Pictures.insert(value("Path"))
              picture.left = This.targetSheet.Range(address).left
              picture.top = This.targetSheet.Range(address).top
              picture.ShapeRange.LockAspectRatio = msoTrue
              'Ensure it fits in cell
              Dim factor As Double: factor = min(cell.height / picture.height, cell.width / picture.width)
              picture.width = picture.width * factor
              'And center picture
              picture.left = cell.left + (cell.width - picture.width) / 2
              picture.top = cell.top + (cell.height - picture.height) / 2
            Case Else
              cell.value = "ERROR: Unsupported dictionary type '" & value("Type") & "'"
          End Select
        Else
          cell.value = "ERROR: Unsupported dictionary"
        End If
      Case Else
        cell.value = "ERROR: Unsupported object type '" & TypeName(value) & "'"
    End Select
  Else
    cell.value = value
  End If
End Sub

Public Function getCell(ByVal address As String) As Object
  Set getCell = This.targetSheet.Range(address)
End Function


'Find all cells in a worksheet that match a pattern
'@param Sheet - The worksheet to search
'@param sPattern - The pattern to search for
'@returns Collection - A collection of cells that match the pattern
Private Function FindAll(ByVal Sheet As Worksheet, ByVal sPattern As String) As Collection
  Set FindAll = New Collection
  Dim v: v = Sheet.UsedRange.value2
  Dim i As Long, j As Long
  For i = 1 To UBound(v, 1)
    For j = 1 To UBound(v, 2)
      If v(i, j) Like sPattern Then
        FindAll.add Sheet.UsedRange.Cells(i, j)
      End If
    Next
  Next
End Function

'Create a dictionary from a list of key-value pairs
'@param v - A list of key-value pairs
'@returns Object<Dictionary> - A dictionary object
Private Function CreateDictionary(ParamArray v() As Variant) As Object
  Set CreateDictionary = CreateObject("Scripting.Dictionary")
  Dim i As Long
  For i = 0 To UBound(v) Step 2
    CreateDictionary.add v(i), v(i + 1)
  Next
End Function

'Find the minimum of two numbers
'@param a - The first number
'@param b - The second number
'@returns - The minimum of the two numbers
Private Function min(ByVal a As Double, ByVal b As Double) As Double
  If a < b Then
    min = a
  Else
    min = b
  End If
End Function





Private Function genIInjector_Create(ByVal templatePath As String, ParamArray v() As Variant) As genIInjector
  Set genIInjector_Create = Create(templatePath, v(0))
End Function
Private Function genIInjector_getFormulaBindings() As stdEnumerator
  Set genIInjector_getFormulaBindings = getFormulaBindings()
End Function
Private Function genIInjector_InitialiseTarget() As Object
  Set genIInjector_InitialiseTarget = InitialiseTarget()
End Function
Private Sub genIInjector_CleanupTarget()
  Call CleanupTarget
End Sub


