VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORMMissingDataRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum EAssetMasterlistType
  SPS
  Overflow
  WWTW
  StormTank
  Syphon
  PipeBridge
  GravitySewer
  RisingMain
  IPS
  Growth
  Infiltration
  H2S
  EAMT_NotApplicable
  HydraulicIncapacity
End Enum

Public Enum EORMDivision
  Infra
  [NonInfra SPS]
  [NonInfra STW]
  [NonInfra IPS]
  [NonInfra Bio]
  [NonInfra Pwg]
  [TBC -Clean]
  EORMDivision_NotApplicable
End Enum

Public Enum EORMMissingDataAction
  [No Action]
  [No action taken - Risk Closed]
  [Manual Review]
  [Clean water review]
  [Manual Review @ 2025]
End Enum

Public Enum EORMFunction
  Waste
  Clean
  Unknown
  EORMFunction_NotApplicable
End Enum

Public Enum EORMNOSStage
  [1. NOS detailed review required (Move to 4B)]
  [2a. TAP investment review required (Move to 5A)]
  [2b. TAP fits existing promoted program. (Move to 5X)]
  [X. OPS to remediate. Capital investment rejected (Move to 10)]
  [Y. OPS Risk deemed resolved (Move to 10)]
End Enum

Private Enum ERowMode
  RowModeCreate = 1
  RowModeUpdate = 2
End Enum

Private Type TStringField
  IncludeInPayload As Boolean
  Value As String
End Type

Private Type TEnumField
  IncludeInPayload As Boolean
  Value As Long
End Type

Private Type TBooleanField
  IncludeInPayload As Boolean
  Value As Boolean
End Type

Private Type TActionsField
  IncludeInPayload As Boolean
  Count As Long
  Values() As EORMMissingDataAction
End Type

Private Type TThis
  Mode As ERowMode
  KeyRiskID As String

  RiskID As TStringField
  AssetType As TEnumField
  AssetID As TStringField
  ORMFunction As TEnumField
  Division As TEnumField
  Comment As TStringField
  NOSStage As TEnumField
  NOSComments As TStringField
  TransferredToSTORM As TBooleanField
  STORMTransferID As TStringField
  Postcode As TStringField
  MultiAsset As TStringField
  Actions As TActionsField
End Type
Private This As TThis

Private Const ORM_LIST_URL As String = "https://contoso.sharepoint.com/sites/OPERATIONALRISKMANAGEMENT/ORM/Lists/ORMMissingData/AllItems.aspx"

'-------------------------
' Constructors
'-------------------------

Public Function CreateCreator(ByVal RiskID As String) As ORMMissingDataRow
  Dim row As ORMMissingDataRow
  Set row = New ORMMissingDataRow
  Call row.protInit(RowModeCreate, RiskID)
  Set CreateCreator = row
End Function

Public Function CreateUpdator(ByVal RiskID As String) As ORMMissingDataRow
  Dim row As ORMMissingDataRow
  Set row = New ORMMissingDataRow
  Call row.protInit(RowModeUpdate, RiskID)
  Set CreateUpdator = row
End Function

Friend Sub protInit(ByVal mode As ERowMode, ByVal RiskID As String)
  RiskID = Trim$(RiskID)
  If LenB(RiskID) = 0 Then
    Err.Raise 5, "ORMMissingDataRow::protInit", "RiskID cannot be blank."
  End If

  This.Mode = mode
  This.KeyRiskID = RiskID
  Erase This.Actions.Values
  This.Actions.Count = 0
  This.Actions.IncludeInPayload = False
End Sub

'-------------------------
' Field properties
'-------------------------

Public Property Get RiskID() As String
  RiskID = This.KeyRiskID
End Property

Public Property Let RiskID(ByVal value As String)
  value = Trim$(value)
  If LenB(value) = 0 Then
    Err.Raise 5, "ORMMissingDataRow::RiskID", "RiskID cannot be blank."
  End If

  This.KeyRiskID = value
  This.RiskID.Value = value
  This.RiskID.IncludeInPayload = True
End Property

Public Property Get AssetType() As EAssetMasterlistType
  AssetType = CLng(This.AssetType.Value)
End Property

Public Property Let AssetType(ByVal value As EAssetMasterlistType)
  Call AssetTypeToText(value)
  This.AssetType.Value = CLng(value)
  This.AssetType.IncludeInPayload = True
End Property

Public Property Get AssetID() As String
  AssetID = This.AssetID.Value
End Property

Public Property Let AssetID(ByVal value As String)
  This.AssetID.Value = value
  This.AssetID.IncludeInPayload = True
End Property

Public Property Get ORMFunction() As EORMFunction
  ORMFunction = CLng(This.ORMFunction.Value)
End Property

Public Property Let ORMFunction(ByVal value As EORMFunction)
  Dim functionText As String
  functionText = ORMFunctionToText(value)
  This.ORMFunction.Value = CLng(value)
  This.ORMFunction.IncludeInPayload = True
End Property

Public Property Get Division() As EORMDivision
  Division = CLng(This.Division.Value)
End Property

Public Property Let Division(ByVal value As EORMDivision)
  Dim divisionText As String
  divisionText = ORMDivisionToText(value)
  This.Division.Value = CLng(value)
  This.Division.IncludeInPayload = True
End Property

Public Property Get Comment() As String
  Comment = This.Comment.Value
End Property

Public Property Let Comment(ByVal value As String)
  This.Comment.Value = value
  This.Comment.IncludeInPayload = True
End Property

Public Property Get NOS_Stage() As EORMNOSStage
  NOS_Stage = CLng(This.NOSStage.Value)
End Property

Public Property Let NOS_Stage(ByVal value As EORMNOSStage)
  Dim nosStageText As String
  nosStageText = ORMNOSStageToText(value)
  This.NOSStage.Value = CLng(value)
  This.NOSStage.IncludeInPayload = True
End Property

Public Property Get NOS_Comments() As String
  NOS_Comments = This.NOSComments.Value
End Property

Public Property Let NOS_Comments(ByVal value As String)
  This.NOSComments.Value = value
  This.NOSComments.IncludeInPayload = True
End Property

Public Property Get TransferredToSTORM() As Boolean
  TransferredToSTORM = This.TransferredToSTORM.Value
End Property

Public Property Let TransferredToSTORM(ByVal value As Boolean)
  This.TransferredToSTORM.Value = value
  This.TransferredToSTORM.IncludeInPayload = True
End Property

Public Property Get STORMTransferID() As String
  STORMTransferID = This.STORMTransferID.Value
End Property

Public Property Let STORMTransferID(ByVal value As String)
  This.STORMTransferID.Value = value
  This.STORMTransferID.IncludeInPayload = True
End Property

Public Property Get Postcode() As String
  Postcode = This.Postcode.Value
End Property

Public Property Let Postcode(ByVal value As String)
  This.Postcode.Value = value
  This.Postcode.IncludeInPayload = True
End Property

Public Property Get MultiAsset() As String
  MultiAsset = This.MultiAsset.Value
End Property

Public Property Let MultiAsset(ByVal value As String)
  This.MultiAsset.Value = value
  This.MultiAsset.IncludeInPayload = True
End Property

Public Function SetActions(ParamArray values() As Variant) As ORMMissingDataRow
  Dim i As Long
  Dim enumValue As EORMMissingDataAction
  Dim hasValues As Boolean

  On Error Resume Next
  hasValues = (UBound(values) >= LBound(values))
  If Err.Number <> 0 Then
    hasValues = False
    Err.Clear
  End If
  On Error GoTo 0

  If Not hasValues Then
    Call ClearActions()
    Set SetActions = Me
    Exit Function
  End If

  ReDim This.Actions.Values(1 To UBound(values) - LBound(values) + 1)
  This.Actions.Count = 0

  For i = LBound(values) To UBound(values)
    enumValue = CLng(values(i))
    Call ORMMissingDataActionToText(enumValue)

    This.Actions.Count = This.Actions.Count + 1
    This.Actions.Values(This.Actions.Count) = enumValue
  Next
  This.Actions.IncludeInPayload = True

  Set SetActions = Me
End Function

Public Function ClearActions() As ORMMissingDataRow
  Erase This.Actions.Values
  This.Actions.Count = 0
  This.Actions.IncludeInPayload = True
  Set ClearActions = Me
End Function

'-------------------------
' Batch execution
'-------------------------

Public Function ExecuteCollection(ByVal rows As Collection) As stdJSON
  Dim payloads As stdJSON
  Dim list As stdSharepointList
  Dim item As Variant
  Dim row As ORMMissingDataRow
  Dim commitPayload As stdJSON

  If rows Is Nothing Then Err.Raise 5, "ORMMissingDataRow::ExecuteCollection", "Rows collection cannot be Nothing."

  Set payloads = stdJSON.Create(eJSONArray)
  Set list = CreateListClient()

  For Each item In rows
    If Not IsObject(item) Then
      Err.Raise 5, "ORMMissingDataRow::ExecuteCollection", "Rows collection must contain ORMMissingDataRow objects."
    End If
    If Not TypeOf item Is ORMMissingDataRow Then
      Err.Raise 5, "ORMMissingDataRow::ExecuteCollection", "Rows collection must contain ORMMissingDataRow objects."
    End If

    Set row = item
    Set commitPayload = row.protBuildCommitPayload(list)
    If Not commitPayload Is Nothing Then payloads.Add commitPayload
  Next

  If payloads.Length = 0 Then
    Set ExecuteCollection = EmptyCommitResult()
    Exit Function
  End If

  Set ExecuteCollection = list.ProcessCommitPayloads(payloads)
End Function

'-------------------------
' Instance payload builders
'-------------------------

Friend Function protBuildCommitPayload(ByVal list As stdSharepointList) As stdJSON
  Dim out As stdJSON
  Dim payload As stdJSON
  Dim data As stdJSON
  Dim itemId As Long

  Select Case This.Mode
    Case RowModeCreate
      Set data = BuildCreateDataPayload()
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "create"
      out.Add "payload", data
      Set protBuildCommitPayload = out

    Case RowModeUpdate
      Set data = BuildUpdateDataPayload()
      If data.Length = 0 Then Exit Function

      itemId = ResolveItemIdByRiskID(list, This.KeyRiskID)
      Set payload = stdJSON.Create(eJSONObject)
      payload.Add "id", itemId
      payload.Add "data", data

      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "update"
      out.Add "payload", payload
      Set protBuildCommitPayload = out

    Case Else
      Err.Raise 5, "ORMMissingDataRow::protBuildCommitPayload", "Unknown row mode."
  End Select
End Function

Private Function BuildCreateDataPayload() As stdJSON
  Dim data As stdJSON

  Set data = stdJSON.Create(eJSONObject)
  data.Add "Title", This.KeyRiskID
  AppendIncludedFields data

  Set BuildCreateDataPayload = data
End Function

Private Function BuildUpdateDataPayload() As stdJSON
  Dim data As stdJSON

  Set data = stdJSON.Create(eJSONObject)
  AppendIncludedFields data

  If This.RiskID.IncludeInPayload Then
    JsonAddOrSet data, "Title", This.RiskID.Value
  End If

  Set BuildUpdateDataPayload = data
End Function

Private Sub AppendIncludedFields(ByVal data As stdJSON)
  If This.AssetType.IncludeInPayload Then JsonAddOrSet data, "AssetType", AssetTypeToText(CLng(This.AssetType.Value))
  If This.AssetID.IncludeInPayload Then JsonAddOrSet data, "AssetID", This.AssetID.Value
  If This.ORMFunction.IncludeInPayload Then JsonAddOrSet data, "Function", ORMFunctionToText(CLng(This.ORMFunction.Value))
  If This.Division.IncludeInPayload Then JsonAddOrSet data, "Division", ORMDivisionToText(CLng(This.Division.Value))
  If This.Comment.IncludeInPayload Then JsonAddOrSet data, "Comment", This.Comment.Value
  If This.NOSStage.IncludeInPayload Then JsonAddOrSet data, "NOS_Stage", ORMNOSStageToText(CLng(This.NOSStage.Value))
  If This.NOSComments.IncludeInPayload Then JsonAddOrSet data, "NOS_Comments", This.NOSComments.Value
  If This.TransferredToSTORM.IncludeInPayload Then JsonAddOrSet data, "TransferredToSTORM", This.TransferredToSTORM.Value
  If This.STORMTransferID.IncludeInPayload Then JsonAddOrSet data, "STORMTransferID", This.STORMTransferID.Value
  If This.Postcode.IncludeInPayload Then JsonAddOrSet data, "Postcode", This.Postcode.Value
  If This.MultiAsset.IncludeInPayload Then JsonAddOrSet data, "MultiAsset", This.MultiAsset.Value
  If This.Actions.IncludeInPayload Then JsonAddOrSet data, "Action", ActionsToJsonArray()
End Sub

Private Function ActionsToJsonArray() As stdJSON
  Dim out As stdJSON
  Dim i As Long

  Set out = stdJSON.Create(eJSONArray)
  For i = 1 To This.Actions.Count
    out.Add ORMMissingDataActionToText(This.Actions.Values(i))
  Next
  Set ActionsToJsonArray = out
End Function

'-------------------------
' SharePoint helpers
'-------------------------

Private Function ResolveItemIdByRiskID(ByVal list As stdSharepointList, ByVal riskID As String) As Long
  Dim inplaceQuery As String
  Dim rows As stdJSON
  Dim candidateRow As stdJSON
  Dim matchedRow As stdJSON
  Dim i As Long

  inplaceQuery = "Title:""" & Replace(riskID, """", "\""") & """"
  Set rows = list.ItemsInplaceQuery(inplaceQuery)
  If rows Is Nothing Then
    Err.Raise 5, "ORMMissingDataRow::ResolveItemIdByRiskID", "Unable to find row for RiskID '" & riskID & "'."
  End If

  If rows.Length = 0 Then
    Err.Raise 5, "ORMMissingDataRow::ResolveItemIdByRiskID", "No row found for RiskID '" & riskID & "'."
  End If

  For i = 1 To rows.Length
    Set candidateRow = rows(i)
    If Not candidateRow Is Nothing Then
      If candidateRow.Exists("Title") Then
        If candidateRow("Title") = riskID Then
          If matchedRow Is Nothing Then
            Set matchedRow = candidateRow
            Exit For
          End If
        End If
      End If
    End If
  Next

  If matchedRow Is Nothing Then
    Err.Raise 5, "ORMMissingDataRow::ResolveItemIdByRiskID", "No exact Title match found for RiskID '" & riskID & "'."
  End If

  If Not matchedRow.Exists("Id") Then
    Err.Raise 5, "ORMMissingDataRow::ResolveItemIdByRiskID", "Matched row for RiskID '" & riskID & "' is missing Id."
  End If

  ResolveItemIdByRiskID = CLng(matchedRow("Id"))
End Function

Private Function CreateListClient() As stdSharepointList
  Dim list As stdSharepointList

  Set list = stdSharepointList.Create(ORM_LIST_URL, ContosoAuth.Authenticator)
  Call list.FieldsAdd("Id", SharePointFieldNumber)
  Call list.FieldsAdd("Title", SharePointFieldText, True)
  Call list.FieldsAdd("AssetType", SharePointFieldChoice, True)
  Call list.FieldsAdd("AssetID", SharePointFieldText, True)
  Call list.FieldsAdd("Function", SharePointFieldChoice, True)
  Call list.FieldsAdd("Division", SharePointFieldChoice, True)
  Call list.FieldsAdd("Comment", SharePointFieldText, True)
  Call list.FieldsAdd("NOS_Stage", SharePointFieldChoice, True)
  Call list.FieldsAdd("NOS_Comments", SharePointFieldText, True)
  Call list.FieldsAdd("TransferredToSTORM", SharePointFieldBoolean, True)
  Call list.FieldsAdd("STORMTransferID", SharePointFieldText, True)
  Call list.FieldsAdd("Postcode", SharePointFieldText, True)
  Call list.FieldsAdd("MultiAsset", SharePointFieldText, True)
  Call list.FieldsAdd("Action", SharePointFieldMultiChoice, True)

  Set CreateListClient = list
End Function

Private Function EmptyCommitResult() As stdJSON
  Dim out As stdJSON
  Set out = stdJSON.Create(eJSONObject)
  out.Add "update", stdJSON.Create(eJSONArray)
  out.Add "create", stdJSON.Create(eJSONArray)
  out.Add "delete", stdJSON.Create(eJSONArray)
  Set EmptyCommitResult = out
End Function

'-------------------------
' Enum conversion helpers
'-------------------------

Private Function AssetTypeToText(ByVal value As EAssetMasterlistType) As String
  Select Case value
    Case SPS: AssetTypeToText = "SPS"
    Case Overflow: AssetTypeToText = "Overflow"
    Case WWTW: AssetTypeToText = "WWTW"
    Case StormTank: AssetTypeToText = "StormTank"
    Case Syphon: AssetTypeToText = "Syphon"
    Case PipeBridge: AssetTypeToText = "PipeBridge"
    Case GravitySewer: AssetTypeToText = "GravitySewer"
    Case RisingMain: AssetTypeToText = "RisingMain"
    Case IPS: AssetTypeToText = "IPS"
    Case Growth: AssetTypeToText = "Growth"
    Case Infiltration: AssetTypeToText = "Infiltration"
    Case H2S: AssetTypeToText = "H2S"
    Case EAMT_NotApplicable: AssetTypeToText = "N/A"
    Case HydraulicIncapacity: AssetTypeToText = "HydraulicIncapacity"
    Case Else
      Err.Raise 5, "ORMMissingDataRow::AssetTypeToText", "Unsupported EAssetMasterlistType value: " & CStr(CLng(value))
  End Select
End Function

Private Function ORMDivisionToText(ByVal value As EORMDivision) As String
  Select Case value
    Case Infra: ORMDivisionToText = "Infra"
    Case [NonInfra SPS]: ORMDivisionToText = "NonInfra SPS"
    Case [NonInfra STW]: ORMDivisionToText = "NonInfra STW"
    Case [NonInfra IPS]: ORMDivisionToText = "NonInfra IPS"
    Case [NonInfra Bio]: ORMDivisionToText = "NonInfra Bio"
    Case [NonInfra Pwg]: ORMDivisionToText = "NonInfra Pwg"
    Case [TBC -Clean]: ORMDivisionToText = "TBC -Clean"
    Case EORMDivision_NotApplicable: ORMDivisionToText = "N/A"
    Case Else
      Err.Raise 5, "ORMMissingDataRow::ORMDivisionToText", "Unsupported EORMDivision value: " & CStr(CLng(value))
  End Select
End Function

Private Function ORMFunctionToText(ByVal value As EORMFunction) As String
  Select Case value
    Case Waste: ORMFunctionToText = "Waste"
    Case Clean: ORMFunctionToText = "Clean"
    Case Unknown: ORMFunctionToText = "Unknown"
    Case EORMFunction_NotApplicable: ORMFunctionToText = "N/A"
    Case Else
      Err.Raise 5, "ORMMissingDataRow::ORMFunctionToText", "Unsupported EORMFunction value: " & CStr(CLng(value))
  End Select
End Function

Private Function ORMNOSStageToText(ByVal value As EORMNOSStage) As String
  Select Case value
    Case [1. NOS detailed review required (Move to 4B)]: ORMNOSStageToText = "1. NOS detailed review required (Move to 4B)"
    Case [2a. TAP investment review required (Move to 5A)]: ORMNOSStageToText = "2a. TAP investment review required (Move to 5A)"
    Case [2b. TAP fits existing promoted program. (Move to 5X)]: ORMNOSStageToText = "2b. TAP fits existing promoted program. (Move to 5X)"
    Case [X. OPS to remediate. Capital investment rejected (Move to 10)]: ORMNOSStageToText = "X. OPS to remediate. Capital investment rejected (Move to 10)"
    Case [Y. OPS Risk deemed resolved (Move to 10)]: ORMNOSStageToText = "Y. OPS Risk deemed resolved (Move to 10)"
    Case Else
      Err.Raise 5, "ORMMissingDataRow::ORMNOSStageToText", "Unsupported EORMNOSStage value: " & CStr(CLng(value))
  End Select
End Function

Private Function ORMMissingDataActionToText(ByVal value As EORMMissingDataAction) As String
  Select Case value
    Case [No Action]: ORMMissingDataActionToText = "No Action"
    Case [No action taken - Risk Closed]: ORMMissingDataActionToText = "No action taken - Risk Closed"
    Case [Manual Review]: ORMMissingDataActionToText = "Manual Review"
    Case [Clean water review]: ORMMissingDataActionToText = "Clean water review"
    Case [Manual Review @ 2025]: ORMMissingDataActionToText = "Manual Review @ 2025"
    Case Else
      Err.Raise 5, "ORMMissingDataRow::ORMMissingDataActionToText", "Unsupported EORMMissingDataAction value: " & CStr(CLng(value))
  End Select
End Function

'-------------------------
' JSON helpers
'-------------------------

Private Sub JsonAddOrSet(ByVal parent As stdJSON, ByVal key As String, ByVal v As Variant)
  If parent Is Nothing Then Exit Sub

  If parent.Exists(key) Then
    If IsObject(v) Then
      Set parent(key) = v
    Else
      parent(key) = v
    End If
  Else
    If IsObject(v) Then
      parent.Add key, v
    Else
      parent.Add key, v
    End If
  End If
End Sub

'-------------------------
' Example usage
'-------------------------

Public Sub ExampleUsage()
  Dim rows As Collection
  Dim createRow As ORMMissingDataRow
  Dim updateRow As ORMMissingDataRow
  Dim result As stdJSON

  Set rows = New Collection

  Set createRow = ORMMissingDataRow.CreateCreator("RISK-123")
  createRow.AssetType = SPS
  createRow.ORMFunction = Waste
  createRow.Division = Infra
  createRow.Comment = "Created from typed updater"
  Call createRow.SetActions([Manual Review], [Clean water review])
  rows.Add createRow

  Set updateRow = ORMMissingDataRow.CreateUpdator("RISK-123")
  updateRow.NOS_Comments = "Reviewed by ORM team"
  updateRow.TransferredToSTORM = True
  rows.Add updateRow

  Set result = ORMMissingDataRow.ExecuteCollection(rows)
  Debug.Print result.ToString()
End Sub
