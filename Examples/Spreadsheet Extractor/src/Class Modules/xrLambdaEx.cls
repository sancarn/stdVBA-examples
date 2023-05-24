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
    
    Set oLib("targetSheet") = Nothing
  End If
  Set GetExtensionLibrary = oLib
End Function

'------------------------------------------------------------------------------------------------------------

'Test function, adds 2 numbers
Public Function add(a As Double, b As Double) As Double
  add = a + b
End Function




'--Testing--
'@test
Friend Sub Test(ByVal sLambda As String)
  With Create(sLambda)
    Set .oFunctExt("targetSheet") = ActiveSheet
    MsgBox .Run()
  End With
End Sub

