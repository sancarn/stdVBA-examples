VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsSrcSharePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements dsISrc

Const SPRoot = "https://company.sharepoint.com"

Private Function dsISrc_getName() As String
  dsISrc_getName = "Sharepoint"
End Function

Private Function dsISrc_linkFiber(ByVal template As stdFiber, ByVal destPath As String, ByVal Data As stdJSON) As stdFiber
  Dim pqFormula As String
  pqFormula = "let Source = SharePoint.Tables(""$siteURL"", [Implementation=null, ApiVersion=15]), Table = Source{[Title=""$listTitle""]}[Items] in Table"
  pqFormula = Replace(pqFormula, "$siteURL", SPRoot & Data.item("Site"))
  pqFormula = Replace(pqFormula, "$listTitle", Data.item("ListTitle"))
  
  Dim json As stdJSON: Set json = stdJSON.Create(eJSONObject)
  json.add "M", pqFormula
  
  'Delegate to dsSrcPowerQuery
  Set dsISrc_linkFiber = dsSrcPowerQuery.AsISrc.linkFiber(template, destPath, json)
End Function

Private Function dsISrc_getDataType() As stdJSON
  Set dsISrc_getDataType = stdJSON.Create(eJSONObject)
  With dsISrc_getDataType
    With .AddObject("Site")
      .add "DisplayText", "Site URL"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    
    With .AddObject("Site")
      .add "DisplayText", "Site URL"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    
    With .AddObject("ListIdentifier")
      .add "DisplayText", "List identifier"
      .add "Tooltip", "How do you want to identify the list?"
      .add "Type", "Dropdown"
      With .AddObject("Dropdown")
        Dim c As New Collection
        c.add "By Title"
        c.add "By ID"
        .add "Options", c
      End With
    End With
    
    With .AddObject("ListTitle")
      .add "DisplayText", "List Title"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
    
    With .AddObject("ListID")
      .add "DisplayText", "List ID"
      .add "Tooltip", ""
      .add "Type", "Text"
    End With
  End With
End Function

Public Function AsISrc() As dsISrc
  Set AsISrc = Me
End Function
