Attribute VB_Name = "mdMain"
Public mdApps As Collection

Public Sub Test()
  Call SaveSetting("RPATest", "RPATest", "A", 0)
  Call SaveSetting("RPATest", "RPATest", "B", 0)
  Call SaveSetting("RPATest", "RPATest", "C", 0)
  Call SaveSetting("RPATest", "RPATest", "D", 0)
  Call SaveSetting("RPATest", "RPATest", "E", 0)
  Call SaveSetting("RPATest", "RPATest", "F", 0)
  Call SaveSetting("RPATest", "RPATest", "G", 0)
  Call SaveSetting("RPATest", "RPATest", "H", 0)
  Call ExecuteAll
  Debug.Print GetSetting("RPATest", "RPATest", "A")
  Debug.Print GetSetting("RPATest", "RPATest", "B")
  Debug.Print GetSetting("RPATest", "RPATest", "C")
  Debug.Print GetSetting("RPATest", "RPATest", "D")
  Debug.Print GetSetting("RPATest", "RPATest", "E")
  Debug.Print GetSetting("RPATest", "RPATest", "F")
  Debug.Print GetSetting("RPATest", "RPATest", "G")
  Debug.Print GetSetting("RPATest", "RPATest", "H")
End Sub

Public Sub ExecuteAll()
  Dim eJobs As stdEnumerator: Set eJobs = stdEnumerator.CreateFromListObject(shJobs.ListObjects("Jobs"))
  Dim job As Object
  For Each job In eJobs.AsCollection
    If job("StatusDate") = Empty Then
      Call setRowCell(job, "Status", "Waiting")
    ElseIf stdLambda.Create(job("Frequency Lambda")).Run(job("StatusDate")) Then
      Call setRowCell(job, "Status", "Waiting")
    End If
  Next
  
  Call Continue(eJobs)
End Sub

Public Sub Continue(Optional eJobs As stdEnumerator)
  If eJobs Is Nothing Then Set eJobs = stdEnumerator.CreateFromListObject(shJobs.ListObjects("Jobs"))
  
  'Convert to mdJob objects
  Dim cJobs As Collection: Set cJobs = New Collection
  Dim dJobs As Object: Set dJobs = CreateObject("Scripting.Dictionary")
  Dim job As Object 'Dictionary<>
  For Each job In eJobs.AsCollection
    Dim oJob As mdJob: Set oJob = mdJob.Create(job("Workbook"), job("Macro"), job("ReadOnly"))
    cJobs.add oJob
    Set dJobs(CStr(job("ID"))) = oJob
    Set oJob.Metadata = job
    Set job("=JOB") = oJob
  Next
  
  'Add dependencies
  For Each job In eJobs.AsCollection
    Dim vDepID
    For Each vDepID In Split(job("Dependencies"), ",")
      Call job("=JOB").protAddDependency(dJobs(vDepID))
    Next
  Next
  
  'Port eJobs from cJobs (collection of mdJob objects)
  Set eJobs = stdEnumerator.CreateFromIEnumVariant(cJobs)
  
  'While jobs exist
  While eJobs.length > 0
    DoEvents
    
    'Find incomplete jobs
    Set eJobs = eJobs.Filter(stdLambda.Create("$1.Status <> ""Complete"" and not $1.Status like ""Error*"""))
    
    'Advance all jobs
    Dim task As mdJob
    For Each task In eJobs.AsCollection
      'Advance progress of task
      task.protStep
      
      'Checking status progresses tasks
      Dim sStatus As String: sStatus = task.Status
      Select Case True
        'If complete, set status and date
        Case sStatus = "Complete"
          Call setRowCell(task.Metadata, "Status", task.Status)
          Call setRowCell(task.Metadata, "StatusDate", Now())
        
        'If error, set status but not date
        Case sStatus Like "Error*"
          Call setRowCell(task.Metadata, "Status", task.Status)
        Case Else
          Call setRowCell(task.Metadata, "Status", task.Status)
      End Select
    Next
  Wend
End Sub



'Given a row object supplied by `stdEnumerator.CreateFromListObject()`, update a cell's value based on it's column name
'@param {Object} Row object
'@param {String} Column name
'@param {Variant} Value to set cell to
Private Sub setRowCell(ByVal oRow As Object, ByVal sColumnName As String, ByVal vValue As Variant)
  Application.Intersect(oRow("=ListRow").Range, oRow("=ListColumns")(sColumnName).Range).value = vValue
  oRow("=ListRow").Range.WrapText = False
End Sub


Sub NewApp()
  Dim app As Application
  Set app = New Application
  app.Visible = False
  app.Workbooks.open "D:\Programming\Github\stdVBA-examples\Examples\MacroDispatcher\Tests\Test.xlsm"
  Call app.OnTime(Now(), "Part7")
  Debug.Assert False
  app.Quit
End Sub
Sub newRuntimeError()
  x = 1 / 0
End Sub
