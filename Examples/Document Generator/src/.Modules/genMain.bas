Attribute VB_Name = "genMain"
#Const DEBUGGING = False

Sub Main()
  'Get Admin lookups
  Dim Lookups As Object: Set Lookups = getLookup(dataAdmin.ListObjects("Lookups"))
  
  'Create factory from target type
  Dim factory As genIInjector
  Select Case Lookups("TargetType")
    Case "Excel"
      Set factory = genInjectorExcel
    Case "PowerPoint"
      Set factory = genInjectorPowerPoint
    Case Else
      Err.Raise 1, "", "No such target type '" & Lookups("TargetType") & "'"
  End Select
  
  'Create injector from factory
  Dim injector As genIInjector
  Set injector = stdLambda.Create(Lookups("TargetLambda")).Run(factory)
  
  'Get bindings
  Dim eBindings As stdEnumerator
  Set eBindings = injector.getFormulaBindings()
  
  'Create lambdas ahead of time
  Dim binding As Object
  For Each binding In eBindings
    Set binding("lambda") = genLambdaEx.Create(binding("lambda"))
  Next
  
  
  Dim SourceFactory As stdLambda: Set SourceFactory = stdLambda.Create(Lookups("Source")).BindGlobal("stdTable", stdTable)
  Dim AfterUpdate As stdLambda: Set AfterUpdate = stdLambda.Create(Lookups("AfterUpdate"))
  
  Dim Source As stdTable: Set Source = SourceFactory.Run()
  Dim rows As stdEnumerator: Set rows = Source.rows
  #If DEBUGGING Then
    Set rows = rows.First(2)
  #End If
  
  Dim row As Object
  For Each row In rows
    'Initialise target (create new document)
    Dim doc As Object: Set doc = injector.InitialiseTarget()
    
    'Evaluate all bindings
    For Each binding In eBindings
      Dim lambdaEx As stdLambda: Set lambdaEx = binding("lambda")
      Dim target As Object: Set target = binding("getSetterTarget").Run()
      Call binding("setter").Run(lambdaEx.Run(row, target))
    Next
    
    'Run post-update lambda
    Call AfterUpdate.Run(doc, row)
    
    'Cleanup the target (close the document etc.)
    injector.CleanupTarget
    
    DoEvents
  Next
End Sub

Private Function getLookup(ByVal lo As ListObject) As Object
  Set getLookup = CreateObject("Scripting.Dictionary")
  Dim v: v = lo.DataBodyRange.value
  For i = 1 To UBound(v, 1)
    getLookup.add v(i, 1), v(i, 2)
  Next
End Function
