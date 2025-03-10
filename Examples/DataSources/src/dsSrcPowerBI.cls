VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcPowerBI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Private Function dsISrc_getName() As String
  dsISrc_getName = "PowerBI"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Dim pqFormula As String
  pqFormula = "let " & _
              "Source = PowerPlatform.Dataflows(null), " & _
              "Workspaces = Source{[Id=""Workspaces""]}[Data], " & _
              "Workspace = Workspaces{[workspaceId=""$WorkspaceID""]}[Data], " & _
              "Dataflow = Workspace{[dataflowId=""$DataFlowID""]}[Data], " & _
              "Dataset = Dataflow{[entity=""$Entity"",version=""""]}[Data] " & _
              "in Dataset"
  pqFormula = Replace(pqFormula, "$WorkspaceID", workspaceID)
  pqFormula = Replace(pqFormula, "$DataFlowID", dataflowID)
  pqFormula = Replace(pqFormula, "$Entity", entityName)
  
  Dim json As stdJSON: Set json = stdJSON.Create(eJSONObject)
  json.add "M", pqFormula
  
  'Delegate to dsSrcPowerQuery
  Set dsISrc_linkFiber = dsSrcPowerQuery.AsISrc.linkFiber(template, destPath, json)
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("WorkspaceID")
      .add "DisplayText", "Workspace ID"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    With .AddObject("DataFlowID")
      .add "DisplayText", "Dataflow ID"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    With .AddObject("Entity")
      .add "DisplayText", "Entity name"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function
