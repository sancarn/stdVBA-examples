VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xrRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private categories As Object

Public Function Create(ByVal lo As ListObject) As xrRules
  Set Create = New xrRules
  Call Create.protInit(lo)
End Function

Friend Sub protInit(lo As ListObject)
  Set categories = stdEnumerator.CreateFromListObject(lo).groupBy(stdLambda.Create("$1.Category"))
  
  Dim vCat, vKey
  For Each vCat In categories.keys()
    Dim oRetDict As Object: Set oRetDict = CreateObject("Scripting.Dictionary")
    Set oRetDict("Path") = xrLambdaEx.Create("targetSheet.parent.fullname")
    Set oRetDict("Sheet") = xrLambdaEx.Create("targetSheet.name")
    Set oRetDict("Category") = xrLambdaEx.Create("targetCategory")
    
    'Apply user defined overrides
    Dim oCat As Object: Set oCat = categories(vCat).item(1)
    For Each vKey In oCat.keys()
      If vKey <> "Category" And Left(vKey, 1) <> "=" Then
        Dim sExpression As String: sExpression = oCat(vKey)
        Set oRetDict(vKey) = xrLambdaEx.Create(sExpression)
      End If
    Next
    
    Set categories(vCat) = oRetDict
  Next
End Sub

'Execute the rules to generate a list of results
'@param {Worksheet} TargetSheet to extract data from
'@param {String} Category to perform which extraction rules
'@returns {stdEnumerator<Dictionary<string, variant>>} Results of extraction as enumerator of dictionaries
Public Function executeRules(ByVal targetSheet As Worksheet, ByVal sCategory As String) As Object
  Dim rules As Object: Set rules = categories(sCategory)
  Dim oRet As Object: Set oRet = CreateObject("Scripting.Dictionary")
  Dim vKey
  For Each vKey In rules.keys()
    Dim ld As stdLambda: Set ld = rules(vKey)
    Set ld.oFunctExt("targetSheet") = targetSheet
    Set ld.oFunctExt("targetCategory") = sCategory
    oRet(vKey) = ld.Run()
  Next
  Set executeRules = oRet
End Function



'Test Rules
'@test
Friend Sub test()
  Dim rules As xrRules
  Set rules = xrRules.Create(dataRules.ListObjects("Rules"))
  Debug.Print stringify(rules.executeRules(testCategories1, "1.0.0"))
  Debug.Print stringify(rules.executeRules(testCategories2, "1.1.0"))
End Sub

'Stringify a dictionary
'@test
'@helper
Private Function stringify(ByVal obj As Object) As String
  Dim s As String: s = "{"
  Dim vKey: For Each vKey In obj.keys()
    s = s & vKey & ": "
    Select Case varType(obj(vKey))
      Case VbVarType.vbString
        s = s & """" & obj(vKey) & """"
      Case Else
        s = s & CDbl(obj(vKey))
    End Select
    s = s & ", "
  Next
  stringify = s & "}"
End Function
