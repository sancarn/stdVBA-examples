VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "genIInjector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Create a new instance of the injector
'@param templatePath - The path to the template file
'@returns - An instance of a genIInjector object - within here include all initialisation stuff
Public Function Create(ByVal templatePath As String, ParamArray v() As Variant) As genIInjector: End Function

'Obtain the bindings for the injector
'@returns stdEnumerator<Object<Dictionary<"lambda", string, "setter", stdICallable>>> - a list of bindings for the injector
'@example ```
'  For each binding in injector.getFormulaBindings()
'    binding("setter").call(stdLambda.create(binding("lambda")).call(row))
'  Next
'```
Public Function getFormulaBindings() As stdEnumerator: End Function

'Initialise the target
Public Function InitialiseTarget() As Object: End Function

'Cleanup the target
Public Sub CleanupTarget(): End Sub
