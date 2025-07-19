VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcGISSTdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Private Const ConnectionString As String = "Provider=sqloledb;Data Source=xxxxxx;Initial Catalog=xxxxxx;User Id=username;Password=password"

Private Function dsISrc_getName() As String
  dsISrc_getName = "GISSTdb"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Dim sql As String: sql = "select *, SHAPE.STAsText() as 'ShapeWKT' from $table"
  sql = Replace(sql, "$table", Data.item("Table"))
  
  Dim json As stdJSON: Set json = stdJSON.Create(eJSONObject)
  With json
    .add "ConnectionString", ConnectionString
    .add "Query", sql
    .add "isBigQuery", Data.item("isBigQuery")
  End With
  
  'Delegate to dsSrcPowerQuery
  Set dsISrc_linkFiber = dsSrcDatabase.AsISrc.linkFiber(template, destPath, json)
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("Table")
      .add "DisplayText", "Table"
      .add "Tooltip", "The table to extract"
      .add "Type", "Text"
    End With
    
    With .AddObject("isBigQuery")
      .add "DisplayText", "Big?"
      .add "Tooltip", "TBC"
      .add "Type", "Boolean"
    End With
  End With
End Function

Public Function createFiberGISST(ByVal fiberTemplate As stdFiber, ByVal gisstLayer As String, ByVal destPath As String) As stdFiber
  Set createFiberGISST = fiberTemplate
  With createFiberGISST
    Const sqlTemplate As String = "select *, SHAPE.STAsText() as 'ShapeWKT' from $table"
    .Meta("SQL") = Replace(sqlTemplate, "$table", gisstLayer)
    .Meta("OutFile") = destPath
    Call .add(stdCallback.CreateFromModule("dsMain", "ProcessGISST").Bind(1), "1. Launch Query")
    Call .add(stdCallback.CreateFromModule("dsMain", "ProcessGISST").Bind(2), "2. Await Completion")
    Call .add(stdCallback.CreateFromModule("dsMain", "ProcessGISST").Bind(3), "3. Export CSV")
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function
