VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Private Function dsISrc_getName() As String
  dsISrc_getName = "DatabaseConnection"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Set dsISrc_linkFiber = template
  With template
    .Meta("ConnectionString") = Data.item("ConnectionString")
    .Meta("Query") = Data.item("Query")
    .Meta("isBigQuery") = Data.item("isBigQuery")
    .Meta("OutPath") = destPath
    
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessDatabase").Bind(1), "1. Launch Query")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessDatabase").Bind(2), "2. Await Completion")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessDatabase").Bind(3), "3. Export CSV")
  End With
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("ConnectionString")
      .add "DisplayText", "Connection String"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    
    With .AddObject("Query")
      .add "DisplayText", "SQL Query"
      .add "Tooltip", ""
      .add "Type", "Multi-Text"
    End With
    
    With .AddObject("isBigQuery")
      .add "DisplayText", "Big?"
      .add "Tooltip", "Is the query going to return over 3k rows? If so tick."
      .add "Type", "Boolean"
    End With
  End With
End Function


'Processes an async GISSTdb OLEdb query.
'@fiberRunner
'@protected
Public Function protProcessDatabase(ByVal stage As Long, ByVal fiber As stdFiber) As Boolean
  With fiber
    Select Case stage
      Case 1 'Launch query
        Set .Meta("db") = DB_Create(.Meta("ConnectionString"))
        Set .Meta("rs") = DB_QueryAsync(.Meta("db"), .Meta("SQL"), .Meta("isBigQuery"))
        protProcessDatabase = True
      Case 2 'Await response from GISST
        protProcessDatabase = Not DB_IsAsyncRunning(.Meta("rs"))
      Case 3 'Dump to CSV
        Call DB_QueryToCSV(.Meta("rs"), .Meta("OutPath"))
        protProcessDatabase = True
    End Select
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function


'Obtain the DB Instance
'@returns Object<ADODB.Connection>
Private Function DB_Create(ByVal ConnectionString As String) As Object
    Const adUseClient As Long = 3
    'Connect to GISSTdb Live server:
    Dim db As Object: Set db = CreateObject("ADODB.connection")
    
    db.CursorLocation = adUseClient 'Allows clone and changing of cursor type
    Call db.open(ConnectionString)
    
    Set DB_Create = db
End Function

'Query DB using SQL
'@param db as Object<ADODB.Connection> - Existing GISSTdb Instance
'@param sSQL - SQL to execute
'@param isLarge - Is query large?
'@returns Object<Adodb.RecordSet> - The recordset queried
Private Function DB_QueryAsync(ByVal db As Object, ByVal sSQL As String, ByVal isLarge As Boolean) As Object
  Const adUseServer = 2, adOpenForwardOnly = 0, adLockReadOnly = 1, adCmdText = 1, adAsyncExecute = &H10
  Dim rs As Object: Set rs = CreateObject("ADODB.RecordSet")
  If isLarge Then rs.CursorLocation = adUseServer
  rs.CacheSize = 1000
  Call rs.open(sSQL, db, adOpenForwardOnly, adLockReadOnly, adCmdText Or adAsyncExecute)
  Set DB_QueryAsync = rs
End Function

'Find out whether an async query is still running.
'@param rs as Object<Adodb.RecordSet> - The record set to check the status of
'@returns - True if query is still runnin, false if query is complete and ready for data extraction
'@example  While DB_IsAsyncRunning(rs): DoEvents: Wend
Private Function DB_IsAsyncRunning(ByVal rs As Object) As Boolean
  Const adStateExecuting = 4
  DB_IsAsyncRunning = (rs.State And adStateExecuting) = adStateExecuting
End Function

'Dump a recordset to a csv
'@param rs as Object<Adodb.RecordSet> - The record set / query to create a csv from
'@param path - the path to the CSV to create
Private Sub DB_QueryToCSV(ByVal rs As Object, ByVal path As String)
  Dim ff As Long: ff = FreeFile()
  Open path For Output As #ff
    Dim iFieldCount As Long: iFieldCount = rs.Fields.count
    Dim sLine As String: sLine = ""
    
    'Load headers into range
    Dim i As Long, j As Long
    For j = 1 To iFieldCount
        If j > 1 Then sLine = sLine & ","
        sLine = sLine & rs.Fields(j - 1).name
    Next
    Print #ff, sLine
    
    'Load data into array
    Do Until rs.EOF
        sLine = ""
        For j = 1 To iFieldCount
            If j > 1 Then sLine = sLine & ","
            sLine = sLine & Serialize(rs.Fields(j - 1).value)
        Next
        Print #ff, sLine
        rs.MoveNext
    Loop
  Close #ff
End Sub

Private Function Serialize(v As Variant) As String
  Select Case VarType(v)
    Case vbString
      Serialize = """" & v & """"
    Case vbDate
      Serialize = """" & Format(v, "yyyy-mm-dd") & """"
    Case vbNull, vbEmpty
      Serialize = ""
    Case Else
      Serialize = v
  End Select
End Function
