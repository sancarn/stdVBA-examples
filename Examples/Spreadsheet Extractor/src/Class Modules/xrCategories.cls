VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xrCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'An array of the following structure
'categories = [
'  {Name: ...  , Conditions: [lambda1, lambda2, ...]}
'  {Name: ...  , Conditions: [lambda1, lambda2, ...]}
']
Private categories As stdArray

Public Function Create(ByVal lo As ListObject) As xrCategories
  Set Create = New xrCategories
  Call Create.protInit(lo)
End Function
Friend Sub protInit(lo As ListObject)
  Set categories = stdArray.Create()
  
  Dim vCat As Object
  For Each vCat In stdEnumerator.CreateFromListObject(lo)
    'Create a new category object, to be used by getCategory
    Dim oNewCategory As Object
    Set oNewCategory = CreateObject("Scripting.Dictionary")
    oNewCategory("Name") = vCat("Category")
     
    'Obtain conditions array, which is an array of lambdas, which can each be used to determine whether a targetSheet matches a category
    Dim conditions As stdArray: Set conditions = stdArray.CreateFromArray(vCat.items())
    Call conditions.Shift 'remove first element
    Set conditions = conditions.Filter(stdLambda.Create("not isObject($1)"))
    
    Set oNewCategory("Conditions") = conditions.Map(stdCallback.CreateFromObjectMethod(xrLambdaEx, "Create"))
    
    'Bind results to array categories global array
    categories.Push oNewCategory
  Next
End Sub

'Obtains a category from a given worksheet
'@param {ByVal Worksheet} Worksheet to identify category within. (VIDEO NOTE: `ByVal` is important here as it crashes without it, due to a VBA bug)
'@returns {String} - "" if no category found, else one of the categories listed in the primal list object
Public Function getCategory(ByVal targetSheet As Worksheet) As String
  For Each Category In categories
    Call Category("Conditions").ForEach(stdLambda.Create("$2.bindGlobal(""targetSheet"",$1)").Bind(targetSheet))
    If Category("Conditions").Reduce(stdLambda.Create("$1 and $2.Run()"), True) Then
      getCategory = Category("Name")
      Exit Function
    End If
  Next
  getCategory = ""
End Function

'Test categories
'@test
Friend Sub test()
  Dim cats As xrCategories
  Set cats = xrCategories.Create(dataCategories.ListObjects("Categories"))
  Debug.Print rules.getCategory(testCategories1)
  Debug.Print cats.getCategory(testCategories2)
End Sub
