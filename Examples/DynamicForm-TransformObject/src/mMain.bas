Attribute VB_Name = "mMain"
Sub TestAlterObject()
  Dim o As DemoClass: Set o = New DemoClass
  o.Test1 = "a"
  o.Test2 = "b"
  
  'Dump input data to sheet
  Call DumpPropsToRange(o, shDemo.Range("A2"))
  
  'Show form and change object
  Call frTransformer.AlterObject(o)
  
  'Dump output data to sheet
  Call DumpPropsToRange(o, shDemo.Range("F2"))
End Sub

Private Sub DumpPropsToRange(obj As Object, r As Range)
  Dim c As Collection: Set c = stdCOM.Create(obj).Properties
  Dim v()
  ReDim v(1 To c.Count, 1 To 2)
  Dim index As Long: index = 0
  Dim prop
  For Each prop In c
    index = index + 1
    v(index, 1) = prop
    v(index, 2) = stdCallback.CreateFromObjectProperty(obj, prop, VbGet)()
  Next
  r.Resize(c.Count, 2).Value = v
End Sub
