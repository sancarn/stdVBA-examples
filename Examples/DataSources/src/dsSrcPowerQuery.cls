VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcPowerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Private Function dsISrc_getName() As String
  dsISrc_getName = "PowerQuery"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Set dsISrc_linkFiber = template
  With dsISrc_linkFiber
    .Meta("OutFile") = destPath
    .Meta("M") = Data.item("M")
    
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessPowerQuery").Bind(1), "1. Creating Query")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessPowerQuery").Bind(2), "2. Executing Query")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessPowerQuery").Bind(3), "3. Awaiting Query Completion")
    Call .add(stdCallback.CreateFromObjectMethod(Me, "protProcessPowerQuery").Bind(4), "4. Exporting to CSV")
  End With
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("M")
      .add "DisplayText", "M Code"
      .add "Tooltip", "The M/Mashup/PowerQuery code to execute."
      .add "Type", "MultiText"
    End With
  End With
End Function

'Processes an excel query.
'@fiberRunner
'@protected
Public Function protProcessPowerQuery(ByVal stage As Long, ByVal fiber As stdFiber) As Boolean
  With fiber
    Select Case stage
      Case 1 'create query of PowerBI data
        Dim xlApp As Excel.Application: Set xlApp = .Agent("xl")
        Dim wb As Workbook: Set wb = xlApp.Workbooks.add()
        Set .Meta("Workbook") = wb
        Set .Meta("Query") = wb.Queries.add("PQ", .Meta("M"))
        protProcessPowerQuery = True
      Case 2 'await completion
        Dim ws As Worksheet: Set ws = .Meta("Workbook").Sheets(1)
        Dim lo As ListObject: Set lo = ws.ListObjects.add(xlSrcQuery, "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=PQ;Extended Properties=""""", Destination:=ws.Range("A1"))
        Set .Meta("ListObject") = lo
        Dim qt As QueryTable: Set qt = lo.QueryTable
        qt.CommandType = xlCmdSql
        qt.CommandText = Array("SELECT * FROM [PQ]")
        Call qt.Refresh(True)
        Set .Meta("QueryTable") = qt
        protProcessPowerQuery = True
      Case 3
        protProcessPowerQuery = Not .Meta("QueryTable").Refreshing
        DoEvents
      Case 4 'perform extraction
        Call RangeToCSV(.Meta("OutFile"), .Meta("ListObject").Range)
        Call .Meta("Workbook").Close(False)
        protProcessPowerQuery = True
    End Select
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function

