VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "genLambdaEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TThis
  meta As Object
End Type
Private This As TThis

'A small regex based DSL compiling to `stdLambda`.
'@param {String} DSL syntax to execute. DSL is superset of stdLambda with addition in remarks.
'@returns {stdLambda} Compiled lambda to execute.
'@remarks Expects global "targetSheet" late-binding. Replaces `r[...]` with `targetSheet.range("...")` and `[...]` with `targetSheet.range("...").value`
Public Function Create(ByVal sLambda As String, Optional ByVal meta As Object = Nothing) As stdLambda
  'Obtain extension library
  Static oFuncExt As Object: If oFuncExt Is Nothing Then Set oFuncExt = GetExtensionLibrary()
  Set This.meta = meta
  
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
    
    'Add extensions
    Dim extension
    For Each extension In stdEnumerator.CreateFromListObject(dataExtensions.ListObjects("Extensions")).AsCollection
      Set oLib(extension("Namespace")) = Application.Run("'" & extension("Path") & "'!getExtension", This.meta)
    Next
  End If
  Set GetExtensionLibrary = oLib
End Function

Public Property Get meta() As Object
  Set meta = This.meta
End Property

'------------------------------------------------------------------------------------------------------------

'Generates a new GUID
'@param {}
'@returns {string} a new random GUID
Public Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Imports an image into a specified range
'@param path - The path to the picture to import
'@returns Object<Dictionary<"Type","Picture","Path",string>> - A "picture obect"
Public Function createPicture(ByVal path As String) As Object
  Set createPicture = CreateObject("Scripting.Dictionary")
  Call createPicture.add("Type", "Picture")
  Call createPicture.add("Path", path)
End Function

''Imports a table into a specified range
''@param table as stdEnumerator<Dictionary<>> -
'Public Function createTable(ByVal table As stdTable) as Object
'  Set createTable = CreateObject("Scripting.Dictionary")
'  Call createTable.add("Type","Table")
'  Call createTable.add("Table", table)
'End Function





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

