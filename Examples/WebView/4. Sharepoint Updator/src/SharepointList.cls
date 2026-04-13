VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharepointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A SharePoint list REST client built on `stdHTTP`, `stdJSON`, and `SharepointAuthenticator`.
'It supports schema-aware field transforms (person/multiperson/multichoice/lookup/date),
'single-item and paged reads, in-place search, item create/update, batch operations, and
'commit-payload orchestration with optional payload simplification.
'@example ```vb
'Dim auth As SharepointAuthenticator
'Set auth = SharepointAuthenticator.Create("https://contoso.sharepoint.com")
'
'Dim risks As SharepointList
'Set risks = SharepointList.Create( _
'  "https://contoso.sharepoint.com/sites/Projects/Lists/Risks/AllItems.aspx", _
'  auth _
')
'
'Call risks.FieldsClear()
'Call risks.FieldsAdd("Title", SharePointFieldText)
'Call risks.FieldsAdd("Owner", SharePointFieldPerson)
'Call risks.FieldsAdd("Tags", SharePointFieldMultiChoice)
'
'Dim rows As stdJSON
'Set rows = risks.ItemsGet()
'Debug.Print rows.Length
'```
'@example ```vb
'Dim payload As stdJSON
'Set payload = stdJSON.Create(eJSONObject)
'Call payload.Add("Title", "New risk")
'Call payload.Add("Owner/Email", "alex@contoso.com")
'
'Dim id As Long
'id = risks.ItemAdd(payload)
'Debug.Print "Created item id:", id
'```
'@example ```vb
'Dim q As stdJSON
'Set q = stdJSON.Create(eJSONObject)
'Call q.Add("$top", 50)
'Call q.Add("$orderby", stdJSON.Create(eJSONArray).AddObject().Add("name", "Modified").Add("direction", "desc"))
'
'Dim paged As stdJSON
'Set paged = risks.ItemsGet(q)
'```

Option Explicit

Public Enum SharePointFieldType
  SharePointFieldText = 0
  SharePointFieldNumber = 1
  SharePointFieldBoolean = 2
  SharePointFieldDate = 3
  SharePointFieldChoice = 4
  SharePointFieldMultiChoice = 5
  SharePointFieldPerson = 6
  SharePointFieldMultiPerson = 7
  SharePointFieldLookup = 8
  SharePointFieldFile = 9
End Enum

Private Type TFieldDef
  InternalName As String
  FieldType As SharePointFieldType
  Nullable As Boolean
  LookupField As String
  SelectProps As Collection
End Type

Private Type TThis
  SiteUrl As String
  ListTitle As String
  ListRelativePath As String
  ListTitleResolved As Boolean

  Authenticator As stdICallable
  Fields() As TFieldDef
  FieldsCount As Long
  UserIdCache As Object

  FormDigestValue As String
  FormDigestExpiresAt As Date
  ListItemEntityType As String
End Type
Private This As TThis

Private Const SP_JSON_VERBOSE As String = "application/json;odata=verbose"
Private Const SP_JSON_LIGHT As String = "application/json"
Private Const BATCH_CHUNK_SIZE As Long = 750
Private Const DIGEST_EXPIRY_BUFFER_SECONDS As Long = 60
Private Const TIMEOUT_MAX_MS As Long = 120000

'Create a SharePoint list client from a SharePoint list URL.
'@constructor
'@param ListUrl - List page URL (for example `/sites/X/Lists/MyList/AllItems.aspx`).
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Variant)=>Void> - Authenticator used by `stdHTTP`.
'@param AutoDiscoverFields - If true, attempts to infer fields from SharePoint metadata.
'@returns - A configured `SharepointList` instance.
Public Function Create( _
  ByVal ListUrl As String, _
  ByVal Authenticator As stdICallable, _
  Optional ByVal AutoDiscoverFields As Boolean = False _
) As SharepointList
  Dim instance As SharepointList
  Set instance = New SharepointList
  Call instance.protInit(ListUrl, Authenticator, AutoDiscoverFields)
  Set Create = instance
End Function

'Initialize list state, schema containers, and caches.
'@protected
'@param ListUrl - List page URL (for example `/sites/X/Lists/MyList/AllItems.aspx`).
'@param Authenticator - Authenticator used by outgoing HTTP requests.
'@param AutoDiscoverFields - If true, attempts to infer fields from SharePoint metadata.
Public Sub protInit( _
  ByVal ListUrl As String, _
  ByVal Authenticator As stdICallable, _
  Optional ByVal AutoDiscoverFields As Boolean = False _
)
  Dim siteUrl As String
  Dim listTitle As String
  Dim relPath As String

  If LenB(Trim$(ListUrl)) = 0 Then
    Err.Raise 5, "SharepointList::protInit", "ListUrl cannot be blank."
  End If
  If Authenticator Is Nothing Then
    Err.Raise 5, "SharepointList::protInit", "Authenticator cannot be Nothing."
  End If

  Call ParseListUrl(ListUrl, siteUrl, listTitle, relPath)

  This.SiteUrl = siteUrl
  This.ListTitle = listTitle
  This.ListRelativePath = relPath
  This.ListTitleResolved = False
  Set This.Authenticator = Authenticator
  This.FieldsCount = 0
  Erase This.Fields
  Set This.UserIdCache = NewDictionary(vbTextCompare)

  This.FormDigestValue = vbNullString
  This.FormDigestExpiresAt = 0
  This.ListItemEntityType = vbNullString

  If AutoDiscoverFields Then
    On Error Resume Next
    Call TryAutoDiscoverFields
    On Error GoTo 0
  End If
End Sub

'Clear all configured schema fields.
'@returns - The same instance for chaining.
Public Function FieldsClear() As SharepointList
  This.FieldsCount = 0
  Erase This.Fields
  Set FieldsClear = Me
End Function

'Add a schema field definition used for transforms and default OData selection.
'@param InternalName - SharePoint internal field name.
'@param FieldType - Field type enum for transform behavior.
'@param Nullable - Whether the field should be treated as nullable.
'@param SelectProps as Variant<Array|Collection|stdJSON> - Optional expanded properties for person/lookup/file fields.
'@param LookupField - Optional lookup target field metadata.
'@returns - The same instance for chaining.
Public Function FieldsAdd( _
  ByVal InternalName As String, _
  ByVal FieldType As SharePointFieldType, _
  Optional ByVal Nullable As Boolean = False, _
  Optional ByVal SelectProps As Variant, _
  Optional ByVal LookupField As String = vbNullString _
) As SharepointList
  Dim key As String
  Dim fieldDef As TFieldDef
  Dim idx As Long

  key = LCase$(Trim$(InternalName))
  If LenB(key) = 0 Then
    Err.Raise 5, "SharepointList::FieldsAdd", "InternalName cannot be blank."
  End If

  fieldDef.InternalName = Trim$(InternalName)
  fieldDef.FieldType = FieldType
  fieldDef.Nullable = CBool(Nullable)
  fieldDef.LookupField = LookupField
  Set fieldDef.SelectProps = NormalizeSelectProps(SelectProps)

  idx = FieldIndexByInternal(key)
  If idx = 0 Then
    This.FieldsCount = This.FieldsCount + 1
    ReDim Preserve This.Fields(1 To This.FieldsCount)
    idx = This.FieldsCount
  End If
  This.Fields(idx) = fieldDef

  Set FieldsAdd = Me
End Function

'Create an equality OData filter object.
'@param FieldName - Field/internal name.
'@param Value - Value to compare.
'@returns as stdJSON<{field: string, operator: "eq", value: Variant}> - OData filter object.
Public Function ODataFilterEq(ByVal FieldName As String, ByVal Value As Variant) As stdJSON
  Set ODataFilterEq = stdJSON.Create(eJSONObject)
  Call ODataFilterEq.Add("field", FieldName)
  Call ODataFilterEq.Add("operator", "eq")
  Call ODataFilterEq.Add("value", Value)
End Function

'Combine filters using logical AND.
'@param filters as ParamArray<stdJSON> - Child filters.
'@returns as stdJSON<{and: stdJSON[]}> - Compound filter object.
Public Function ODataFilterAnd(ParamArray filters() As Variant) As stdJSON
  Dim arr As stdJSON
  Dim i As Long

  Set ODataFilterAnd = stdJSON.Create(eJSONObject)
  Set arr = stdJSON.Create(eJSONArray)
  For i = LBound(filters) To UBound(filters)
    If IsObject(filters(i)) Then
      If TypeOf filters(i) Is stdJSON Then arr.Add filters(i)
    End If
  Next
  ODataFilterAnd.Add "and", arr
End Function

'Combine filters using logical OR.
'@param filters as ParamArray<stdJSON> - Child filters.
'@returns as stdJSON<{or: stdJSON[]}> - Compound filter object.
Public Function ODataFilterOr(ParamArray filters() As Variant) As stdJSON
  Dim arr As stdJSON
  Dim i As Long

  Set ODataFilterOr = stdJSON.Create(eJSONObject)
  Set arr = stdJSON.Create(eJSONArray)
  For i = LBound(filters) To UBound(filters)
    If IsObject(filters(i)) Then
      If TypeOf filters(i) Is stdJSON Then arr.Add filters(i)
    End If
  Next
  ODataFilterOr.Add "or", arr
End Function

'Create an OData order-by descriptor.
'@param FieldName - Field/internal name.
'@param Desc - If true, direction is `desc`; otherwise `asc`.
'@returns as stdJSON<{name: string, direction: "asc"|"desc"}> - Order-by descriptor.
Public Function ODataOrderBy(ByVal FieldName As String, Optional ByVal Desc As Boolean = False) As stdJSON
  Set ODataOrderBy = stdJSON.Create(eJSONObject)
  Call ODataOrderBy.Add("name", FieldName)
  Call ODataOrderBy.Add("direction", IIf(Desc, "desc", "asc"))
End Function

'Get one list item by id.
'@param ItemId - SharePoint item id.
'@param Query as Variant<String|stdJSON> - Optional OData query (raw or structured).
'@returns as stdJSON - Item data transformed to client shape.
Public Function ItemGet(ByVal ItemId As Variant, Optional ByVal Query As Variant) As stdJSON
  Dim api As String
  Dim data As stdJSON
  Dim d As stdJSON

  api = ListItemsEndpoint("(" & CStr(ItemId) & ")" & ODataQueryToString(Query))
  Set data = RequestJson(api, "GET")
  Set d = JsonGetObject(data, "d")
  Set ItemGet = TransformGetterItem(d)
End Function

'Get list items, following SharePoint pagination via `__next`.
'@param Query as Variant<String|stdJSON> - Optional OData query (raw or structured).
'@remark If you have not indexed your columns, and the list has more than 5000 items, 
'you will need to use `ItemsInplaceQuery` instead. Otherwise you will run into the 
'"The attempted operation is prohibited because it exceeds the list view threshold." error
'@returns as stdJSON<Array> - Array of transformed list items.
Public Function ItemsGet(Optional ByVal Query As Variant) As stdJSON
  Dim outArr As stdJSON
  Dim nextUrl As String
  Dim pageJson As stdJSON
  Dim d As stdJSON
  Dim rows As stdJSON
  Dim i As Long

  Set outArr = stdJSON.Create(eJSONArray)
  nextUrl = ListItemsEndpoint(ODataQueryToString(Query))

  Do While LenB(nextUrl) > 0
    Set pageJson = RequestJson(nextUrl, "GET")
    Set d = JsonGetObject(pageJson, "d")

    If d Is Nothing Then Exit Do
    If d.Exists("results") And IsObject(d("results")) And TypeOf d("results") Is stdJSON Then
      Set rows = d("results")
      For i = 1 To rows.Length
        outArr.Add TransformGetterItem(rows(i))
      Next
    End If

    nextUrl = vbNullString
    If d.Exists("__next") Then nextUrl = CStr(d("__next"))
  Loop

  Set ItemsGet = outArr
End Function

'Run SharePoint inplace query via `RenderListDataAsStream`.
'@param InplaceQuery - SharePoint inplace query value for `InplaceSearchQuery`.
'@returns as stdJSON<Array> - Array of transformed list items.
Public Function ItemsInplaceQuery(ByVal InplaceQuery As String) As stdJSON
  Dim api As String
  Dim result As stdJSON
  Dim rows As stdJSON
  Dim outArr As stdJSON
  Dim i As Long

  api = ListEndpoint("/RenderListDataAsStream?InplaceSearchQuery=" & stdHTTP.URIEncode(InplaceQuery))
  Set result = RequestJson(api, "POST", "{}", Empty, SP_JSON_VERBOSE)

  Set outArr = stdJSON.Create(eJSONArray)
  If result.Exists("Row") And IsObject(result("Row")) And TypeOf result("Row") Is stdJSON Then
    Set rows = result("Row")
    For i = 1 To rows.Length
      outArr.Add TransformGetterItem(rows(i))
    Next
  End If
  Set ItemsInplaceQuery = outArr
End Function

'Create a list item.
'@param Data as stdJSON<Object> - Client payload object.
'@param ItemType - Optional list item entity type.
'@returns - Newly created SharePoint item id.
Public Function ItemAdd(ByVal Data As stdJSON, Optional ByVal ItemType As String = vbNullString) As Long
  Dim postData As stdJSON
  Dim response As stdJSON
  Dim d As stdJSON

  If Data Is Nothing Then Err.Raise 5, "SharepointList::ItemAdd", "Data cannot be Nothing."

  Set postData = TransformSetterData(Data)
  If LenB(ItemType) = 0 Then ItemType = GetListItemEntityType()
  SetMetadataType postData, ItemType

  Set response = RequestJson(ListItemsEndpoint(vbNullString), "POST", postData.ToString(), Empty, SP_JSON_VERBOSE)
  Set d = JsonGetObject(response, "d")
  If d Is Nothing Or Not d.Exists("Id") Then
    Err.Raise 5, "SharepointList::ItemAdd", "SharePoint did not return created item id."
  End If
  ItemAdd = CLng(d("Id"))
End Function

'Update (MERGE) a list item.
'@param ItemId - SharePoint item id.
'@param Data as stdJSON<Object> - Partial item payload.
'@returns as stdJSON - Response summary (`status`, `statusText`, optional `data`).
Public Function ItemSet(ByVal ItemId As Variant, ByVal Data As stdJSON) As stdJSON
  Dim patchData As stdJSON
  Dim headers As Variant
  Dim req As stdHTTP
  Dim out As stdJSON

  If Data Is Nothing Then Err.Raise 5, "SharepointList::ItemSet", "Data cannot be Nothing."

  Set patchData = TransformSetterData(Data)
  SetMetadataType patchData, GetListItemEntityType()

  headers = BuildDefaultHeaders(True, SP_JSON_VERBOSE)
  stdHTTP.SetHeader headers, "IF-MATCH", "*"
  stdHTTP.SetHeader headers, "X-HTTP-Method", "MERGE"

  Set req = SendRequest(ListItemsEndpoint("(" & CStr(ItemId) & ")"), "POST", patchData.ToString(), headers, SP_JSON_VERBOSE)
  Set out = stdJSON.Create(eJSONObject)
  out.Add "status", req.ResponseStatus
  out.Add "statusText", req.ResponseStatusText
  If LenB(CStr(req.ResponseText)) > 0 Then
    On Error Resume Next
    out.Add "data", stdJSON.CreateFromString(req.ResponseText)
    On Error GoTo 0
  End If
  Set ItemSet = out
End Function

'Resolve/ensure a SharePoint user id from login/email text.
'@param UserText - User login, email, or principal text.
'@returns - SharePoint user id.
Public Function GetOrCreateUserId(ByVal UserText As String) As Long
  Dim key As String
  Dim body As String
  Dim json As stdJSON
  Dim d As stdJSON

  key = LCase$(Trim$(UserText))
  If LenB(key) = 0 Then Err.Raise 5, "SharepointList::GetOrCreateUserId", "UserText cannot be blank."

  If This.UserIdCache.Exists(key) Then
    GetOrCreateUserId = CLng(This.UserIdCache(key))
    Exit Function
  End If

  body = "{""logonName"":" & JsonStringLiteral(UserText) & "}"
  Set json = RequestJson(This.SiteUrl & "/_api/web/ensureuser", "POST", body, Empty, SP_JSON_VERBOSE)
  Set d = JsonGetObject(json, "d")
  If d Is Nothing Or Not d.Exists("Id") Then
    Err.Raise 5, "SharepointList::GetOrCreateUserId", "SharePoint ensureuser did not return Id."
  End If

  This.UserIdCache.Add key, CLng(d("Id"))
  GetOrCreateUserId = CLng(d("Id"))
End Function

'Batch update list items.
'@param BatchChanges as stdJSON<Array<{id: long, data: stdJSON}>> - Update operations.
'@param ChangeSetId - Optional changeset guid.
'@param BatchUuid - Optional batch guid.
'@returns as stdJSON<Array> - Per-item result objects.
Public Function BatchItemsSet( _
  ByVal BatchChanges As stdJSON, _
  Optional ByVal ChangeSetId As String = vbNullString, _
  Optional ByVal BatchUuid As String = vbNullString _
) As stdJSON
  Dim chunked As Collection
  Dim chunk As stdJSON
  Dim i As Long
  Dim lineItems As Collection
  Dim results As stdJSON
  Dim parseRows As Collection
  Dim digest As String
  Dim headers As Variant
  Dim req As stdHTTP
  Dim itemType As String
  Dim item As stdJSON
  Dim op As stdJSON
  Dim j As Long
  Dim one As Object
  Dim out As stdJSON

  Set results = stdJSON.Create(eJSONArray)
  If BatchChanges Is Nothing Or BatchChanges.Length = 0 Then
    Set BatchItemsSet = results
    Exit Function
  End If

  If LenB(ChangeSetId) = 0 Then ChangeSetId = GenerateUuid()
  If LenB(BatchUuid) = 0 Then BatchUuid = GenerateUuid()
  digest = GetFormDigest()
  itemType = GetListItemEntityType()

  Set chunked = SplitJsonArrayIntoChunks(BatchChanges, BATCH_CHUNK_SIZE)
  For i = 1 To chunked.Count
    Set chunk = chunked(i)
    If chunk.Length > 0 Then chunk.Add chunk(1) 'SP workaround parity with TypeScript implementation

    Set lineItems = New Collection
    lineItems.Add "--batch_" & BatchUuid
    lineItems.Add "Content-Type: multipart/mixed; boundary=changeset_" & ChangeSetId
    lineItems.Add vbNullString

    For j = 1 To chunk.Length
      Set item = chunk(j)
      Set op = TransformSetterData(JsonGetObject(item, "data"))
      SetMetadataType op, itemType

      lineItems.Add "--changeset_" & ChangeSetId
      lineItems.Add "Content-Type: application/http"
      lineItems.Add "Content-Transfer-Encoding: binary"
      lineItems.Add vbNullString
      lineItems.Add "PATCH " & ListItemsEndpoint("(" & CStr(item("id")) & ")") & " HTTP/1.1"
      lineItems.Add "Content-Type: " & SP_JSON_VERBOSE
      lineItems.Add "Accept: " & SP_JSON_LIGHT
      lineItems.Add "If-Match: *"
      lineItems.Add "X-HTTP-Method: MERGE"
      lineItems.Add vbNullString
      lineItems.Add op.ToString()
      lineItems.Add vbNullString
    Next

    lineItems.Add "--changeset_" & ChangeSetId & "--"
    lineItems.Add "--batch_" & BatchUuid & "--"

    headers = BuildDefaultHeaders(True, "multipart/mixed; boundary=""batch_" & BatchUuid & """")
    stdHTTP.SetHeader headers, "X-RequestDigest", digest
    Set req = SendRequest(This.SiteUrl & "/_api/$batch", "POST", JoinCollection(lineItems, vbCrLf), headers, "multipart/mixed")

    Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
    For j = 1 To parseRows.Count - 1
      Set one = parseRows(j)
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "update"
      out.Add "id", chunk(j)("id")
      out.Add "httpStatus", one("statusCode")
      If IsObject(one("json")) Then out.Add "data", one("json")
      out.Add "payload", chunk(j)
      results.Add out
    Next
  Next

  Set BatchItemsSet = results
End Function

'Batch create list items.
'@param BatchItems as stdJSON<Array<stdJSON>> - Create payloads.
'@param BatchUuid - Optional batch guid.
'@returns as stdJSON<Array> - Per-item result objects.
Public Function BatchItemsCreate(ByVal BatchItems As stdJSON, Optional ByVal BatchUuid As String = vbNullString) As stdJSON
  Dim chunked As Collection
  Dim chunk As stdJSON
  Dim i As Long
  Dim j As Long
  Dim lineItems As Collection
  Dim results As stdJSON
  Dim parseRows As Collection
  Dim digest As String
  Dim headers As Variant
  Dim req As stdHTTP
  Dim itemType As String
  Dim op As stdJSON
  Dim one As Object
  Dim out As stdJSON
  Dim d As stdJSON

  Set results = stdJSON.Create(eJSONArray)
  If BatchItems Is Nothing Or BatchItems.Length = 0 Then
    Set BatchItemsCreate = results
    Exit Function
  End If

  If LenB(BatchUuid) = 0 Then BatchUuid = GenerateUuid()
  digest = GetFormDigest()
  itemType = GetListItemEntityType()

  Set chunked = SplitJsonArrayIntoChunks(BatchItems, BATCH_CHUNK_SIZE)
  For i = 1 To chunked.Count
    Set chunk = chunked(i)

    Set lineItems = New Collection
    lineItems.Add "--batch_" & BatchUuid
    lineItems.Add "Content-Type: multipart/mixed; boundary=changeset_" & BatchUuid
    lineItems.Add vbNullString

    For j = 1 To chunk.Length
      Set op = TransformSetterData(chunk(j))
      SetMetadataType op, itemType

      lineItems.Add "--changeset_" & BatchUuid
      lineItems.Add "Content-Type: application/http"
      lineItems.Add "Content-Transfer-Encoding: binary"
      lineItems.Add vbNullString
      lineItems.Add "POST " & ListItemsEndpoint(vbNullString) & " HTTP/1.1"
      lineItems.Add "Content-Type: " & SP_JSON_VERBOSE
      lineItems.Add "Accept: " & SP_JSON_VERBOSE
      lineItems.Add vbNullString
      lineItems.Add op.ToString()
      lineItems.Add vbNullString
    Next

    lineItems.Add "--changeset_" & BatchUuid & "--"
    lineItems.Add "--batch_" & BatchUuid & "--"

    headers = BuildDefaultHeaders(True, "multipart/mixed; boundary=""batch_" & BatchUuid & """")
    stdHTTP.SetHeader headers, "X-RequestDigest", digest
    Set req = SendRequest(This.SiteUrl & "/_api/$batch", "POST", JoinCollection(lineItems, vbCrLf), headers, "multipart/mixed")

    Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
    For j = 1 To parseRows.Count
      Set one = parseRows(j)
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "create"
      out.Add "httpStatus", one("statusCode")
      If IsObject(one("json")) Then
        out.Add "data", one("json")
        Set d = JsonGetObject(one("json"), "d")
        If Not d Is Nothing And d.Exists("Id") Then out.Add "id", CLng(d("Id"))
      End If
      out.Add "payload", chunk(j)
      results.Add out
    Next
  Next

  Set BatchItemsCreate = results
End Function

'Batch delete list items.
'@param ItemIds as stdJSON<Array<long>> - Item ids to delete.
'@param ChangeSetId - Optional changeset guid.
'@param BatchUuid - Optional batch guid.
'@returns as stdJSON<Array> - Per-item result objects.
Public Function BatchItemsDelete( _
  ByVal ItemIds As stdJSON, _
  Optional ByVal ChangeSetId As String = vbNullString, _
  Optional ByVal BatchUuid As String = vbNullString _
) As stdJSON
  Dim chunked As Collection
  Dim chunk As stdJSON
  Dim i As Long
  Dim j As Long
  Dim lineItems As Collection
  Dim results As stdJSON
  Dim parseRows As Collection
  Dim digest As String
  Dim headers As Variant
  Dim req As stdHTTP
  Dim one As Object
  Dim out As stdJSON

  Set results = stdJSON.Create(eJSONArray)
  If ItemIds Is Nothing Or ItemIds.Length = 0 Then
    Set BatchItemsDelete = results
    Exit Function
  End If

  If LenB(ChangeSetId) = 0 Then ChangeSetId = GenerateUuid()
  If LenB(BatchUuid) = 0 Then BatchUuid = GenerateUuid()
  digest = GetFormDigest()

  Set chunked = SplitJsonArrayIntoChunks(ItemIds, BATCH_CHUNK_SIZE)
  For i = 1 To chunked.Count
    Set chunk = chunked(i)

    Set lineItems = New Collection
    lineItems.Add "--batch_" & BatchUuid
    lineItems.Add "Content-Type: multipart/mixed; boundary=changeset_" & ChangeSetId
    lineItems.Add vbNullString

    For j = 1 To chunk.Length
      lineItems.Add "--changeset_" & ChangeSetId
      lineItems.Add "Content-Type: application/http"
      lineItems.Add "Content-Transfer-Encoding: binary"
      lineItems.Add vbNullString
      lineItems.Add "DELETE " & ListItemsEndpoint("(" & CStr(chunk(j)) & ")") & " HTTP/1.1"
      lineItems.Add "Accept: " & SP_JSON_VERBOSE
      lineItems.Add "If-Match: *"
      lineItems.Add vbNullString
    Next

    lineItems.Add "--changeset_" & ChangeSetId & "--"
    lineItems.Add "--batch_" & BatchUuid & "--"

    headers = BuildDefaultHeaders(True, "multipart/mixed; boundary=""batch_" & BatchUuid & """")
    stdHTTP.SetHeader headers, "X-RequestDigest", digest
    Set req = SendRequest(This.SiteUrl & "/_api/$batch", "POST", JoinCollection(lineItems, vbCrLf), headers, "multipart/mixed")

    Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
    For j = 1 To parseRows.Count
      Set one = parseRows(j)
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "delete"
      out.Add "id", chunk(j)
      out.Add "httpStatus", one("statusCode")
      If IsObject(one("json")) Then out.Add "data", one("json")
      results.Add out
    Next
  Next

  Set BatchItemsDelete = results
End Function

'Simplify a commit payload by removing unchanged update fields.
'@param Payload as stdJSON - Commit payload (`create`/`update`/`delete`).
'@returns - Simplified payload, or `Nothing` when an update has no effective field changes.
'@remark For updates, pass `previousData` on the payload object to enable simplification.
Public Function SimplifyPayload(ByVal Payload As stdJSON) As stdJSON
  Dim pType As String
  Dim dataObj As stdJSON
  Dim prevObj As stdJSON
  Dim changed As stdJSON
  Dim info As Collection
  Dim i As Long
  Dim node As Object
  Dim key As String
  Dim curVal As Variant
  Dim prevVal As Variant
  Dim out As stdJSON
  Dim payloadObj As stdJSON

  If Payload Is Nothing Then Exit Function
  If Not Payload.Exists("type") Then
    Set SimplifyPayload = JsonClone(Payload)
    Exit Function
  End If

  pType = LCase$(CStr(Payload("type")))
  If pType <> "update" Then
    Set SimplifyPayload = JsonClone(Payload)
    Exit Function
  End If

  Set payloadObj = JsonGetObject(Payload, "payload")
  Set dataObj = JsonGetObject(payloadObj, "data")
  Set prevObj = JsonGetObject(Payload, "previousData")
  If dataObj Is Nothing Or prevObj Is Nothing Then
    Set SimplifyPayload = JsonClone(Payload)
    Exit Function
  End If

  Set changed = stdJSON.Create(eJSONObject)
  Set info = dataObj.ChildrenInfo
  For i = 1 To info.Count
    Set node = info(i)
    key = CStr(node("key"))
    curVal = node("value")

    If prevObj.Exists(key) Then
      prevVal = prevObj(key)
      If Not ValuesEqualForField(key, curVal, prevVal) Then JsonAddOrSet changed, key, curVal
    Else
      JsonAddOrSet changed, key, curVal
    End If
  Next

  If changed.Length = 0 Then Exit Function

  Set out = stdJSON.Create(eJSONObject)
  out.Add "type", "update"
  Set payloadObj = stdJSON.Create(eJSONObject)
  payloadObj.Add "id", Payload("payload")("id")
  payloadObj.Add "data", changed
  out.Add "payload", payloadObj

  Set SimplifyPayload = out
End Function

'Process a set of commit payloads through batch methods.
'@param Payloads as stdJSON<Array> - Commit payloads with `type` + `payload`.
'@returns as stdJSON<{update: stdJSON[], create: stdJSON[], delete: stdJSON[]}> - Batch operation results.
Public Function ProcessCommitPayloads(ByVal Payloads As stdJSON) As stdJSON
  Dim updates As stdJSON
  Dim creates As stdJSON
  Dim deletes As stdJSON
  Dim simplified As stdJSON
  Dim i As Long
  Dim payload As stdJSON
  Dim pType As String
  Dim out As stdJSON
  Dim updResult As stdJSON
  Dim crtResult As stdJSON
  Dim delResult As stdJSON

  Set updates = stdJSON.Create(eJSONArray)
  Set creates = stdJSON.Create(eJSONArray)
  Set deletes = stdJSON.Create(eJSONArray)

  If Not Payloads Is Nothing Then
    For i = 1 To Payloads.Length
      Set payload = Payloads(i)
      If payload Is Nothing Then GoTo ContinueLoop
      If Not payload.Exists("type") Then GoTo ContinueLoop

      pType = LCase$(CStr(payload("type")))
      Select Case pType
        Case "update"
          Set simplified = SimplifyPayload(payload)
          If Not simplified Is Nothing Then updates.Add simplified("payload")
        Case "create"
          creates.Add payload("payload")
        Case "delete"
          If IsObject(payload("payload")) And TypeOf payload("payload") Is stdJSON Then
            If payload("payload").Exists("id") Then deletes.Add payload("payload")("id")
          Else
            deletes.Add payload("payload")
          End If
      End Select
ContinueLoop:
    Next
  End If

  If updates.Length > 0 Then
    Set updResult = BatchItemsSet(updates)
  Else
    Set updResult = stdJSON.Create(eJSONArray)
  End If

  If creates.Length > 0 Then
    Set crtResult = BatchItemsCreate(creates)
  Else
    Set crtResult = stdJSON.Create(eJSONArray)
  End If

  If deletes.Length > 0 Then
    Set delResult = BatchItemsDelete(deletes)
  Else
    Set delResult = stdJSON.Create(eJSONArray)
  End If

  Set out = stdJSON.Create(eJSONObject)
  out.Add "update", updResult
  out.Add "create", crtResult
  out.Add "delete", delResult

  Set ProcessCommitPayloads = out
End Function

'---------------------------------------
' Internal helpers
'---------------------------------------

 'Build the list items endpoint using a suffix segment/query.
Private Function ListItemsEndpoint(ByVal suffix As String) As String
  ListItemsEndpoint = ListEndpoint("/items" & suffix)
End Function

 'Build the base list REST endpoint using list title and suffix.
Private Function ListEndpoint(ByVal suffix As String) As String
  EnsureListDisplayTitle
  ListEndpoint = This.SiteUrl & "/_api/web/lists/GetByTitle('" & EscapeODataLiteral(This.ListTitle) & "')" & suffix
End Function

 'Send an HTTP request and parse JSON response (empty body -> empty object).
Private Function RequestJson( _
  ByVal url As String, _
  ByVal method As String, _
  Optional ByVal body As String = vbNullString, _
  Optional ByVal extraHeaders As Variant, _
  Optional ByVal contentType As String = SP_JSON_VERBOSE _
) As stdJSON
  Dim req As stdHTTP
  Dim txt As String
  Dim parsed As stdJSON

  Set req = SendRequest(url, method, body, extraHeaders, contentType)
  txt = CStr(req.ResponseText)
  If LenB(txt) = 0 Then
    Set RequestJson = stdJSON.Create(eJSONObject)
    Exit Function
  End If
  Set parsed = stdJSON.CreateFromString(txt)
  Set RequestJson = parsed
End Function

 'Create and send an authenticated SharePoint HTTP request.
Private Function SendRequest( _
  ByVal url As String, _
  ByVal method As String, _
  Optional ByVal body As String = vbNullString, _
  Optional ByVal extraHeaders As Variant, _
  Optional ByVal contentType As String = SP_JSON_VERBOSE _
) As stdHTTP
  Dim headers As Variant
  Dim options As Variant
  Dim req As stdHTTP

  headers = BuildDefaultHeaders(True, contentType)
  MergeHeaderArray headers, extraHeaders

  Set options = stdHTTP.CreateOptions(body, headers, timeoutMS:=TIMEOUT_MAX_MS )

  Set req = stdHTTP.Create(url, method, HTTPSync, options, This.Authenticator)
  EnsureSuccess req, "SharepointList::SendRequest"
  Set SendRequest = req
End Function

 'Build default request headers with optional Accept header.
Private Function BuildDefaultHeaders(ByVal includeAccept As Boolean, ByVal contentType As String) As Variant
  If includeAccept Then
    BuildDefaultHeaders = stdHTTP.CreateHeaders("Accept", SP_JSON_VERBOSE, "Content-Type", contentType)
  Else
    BuildDefaultHeaders = stdHTTP.CreateHeaders("Content-Type", contentType)
  End If
End Function

 'Merge header pairs from one header-array into another.
Private Sub MergeHeaderArray(ByRef headers As Variant, ByVal extraHeaders As Variant)
  Dim i As Long
  If Not IsArray(extraHeaders) Then Exit Sub
  For i = LBound(extraHeaders) To UBound(extraHeaders) Step 2
    stdHTTP.SetHeader headers, CStr(extraHeaders(i)), CStr(extraHeaders(i + 1))
  Next
End Sub

 'Raise an error when the HTTP response status is non-success.
Private Sub EnsureSuccess(ByVal req As stdHTTP, ByVal sourceName As String)
  Dim code As Long
  Dim details As String

  code = CLng(req.ResponseStatus)
  If code >= 200 And code < 300 Then Exit Sub

  details = CStr(req.ResponseText)
  If LenB(details) > 800 Then details = Left$(details, 800) & "..."
  Err.Raise vbObjectError + 513, sourceName, "HTTP " & CStr(code) & " " & CStr(req.ResponseStatusText) & IIf(LenB(details) > 0, " | " & details, vbNullString)
End Sub

 'Convert structured/raw OData query input into a querystring.
Private Function ODataQueryToString(Optional ByVal Query As Variant) As String
  Dim params As Collection
  Dim q As stdJSON
  Dim selectCsv As String
  Dim expandCsv As String
  Dim filterText As String
  Dim useDefaultExpand As Boolean

  Set params = New Collection
  useDefaultExpand = True

  If IsMissingOrEmpty(Query) Then
    BuildDefaultSelectExpand selectCsv, expandCsv
    If LenB(selectCsv) > 0 Then params.Add "$select=" & selectCsv
    If LenB(expandCsv) > 0 Then params.Add "$expand=" & expandCsv
  ElseIf VarType(Query) = vbString Then
    ODataQueryToString = NormalizeRawQuery(CStr(Query))
    Exit Function
  ElseIf IsObject(Query) And TypeOf Query Is stdJSON Then
    Set q = Query
    If q.Exists("$select") Then
      If IsObject(q("$select")) And TypeOf q("$select") Is stdJSON Then
        selectCsv = EncodeArrayCsv(q("$select"))
        If LenB(selectCsv) > 0 Then params.Add "$select=" & selectCsv
        useDefaultExpand = False
      ElseIf LCase$(Trim$(CStr(q("$select")))) = "*" Then
        BuildDefaultSelectExpand selectCsv, expandCsv
        If LenB(selectCsv) > 0 Then params.Add "$select=" & selectCsv
      End If
    Else
      BuildDefaultSelectExpand selectCsv, expandCsv
      If LenB(selectCsv) > 0 Then params.Add "$select=" & selectCsv
    End If

    If q.Exists("$expand") Then
      If IsObject(q("$expand")) And TypeOf q("$expand") Is stdJSON Then
        expandCsv = EncodeArrayCsv(q("$expand"))
        If LenB(expandCsv) > 0 Then params.Add "$expand=" & expandCsv
      End If
      useDefaultExpand = False
    ElseIf useDefaultExpand Then
      BuildDefaultSelectExpand selectCsv, expandCsv
      If LenB(expandCsv) > 0 Then params.Add "$expand=" & expandCsv
    End If

    If q.Exists("$top") Then params.Add "$top=" & CStr(q("$top"))
    If q.Exists("$skip") Then params.Add "$skip=" & CStr(q("$skip"))
    If q.Exists("$inlinecount") Then params.Add "$inlinecount=" & stdHTTP.URIEncode(CStr(q("$inlinecount")))
    If q.Exists("$format") Then params.Add "$format=" & stdHTTP.URIEncode(CStr(q("$format")))
    If q.Exists("$selectversion") Then
      If CBool(q("$selectversion")) Then params.Add "$selectversion=true"
    End If

    If q.Exists("$orderby") And IsObject(q("$orderby")) And TypeOf q("$orderby") Is stdJSON Then
      params.Add "$orderby=" & BuildOrderByString(q("$orderby"))
    End If

    If q.Exists("$filter") And IsObject(q("$filter")) And TypeOf q("$filter") Is stdJSON Then
      filterText = ODataFilterToString(q("$filter"))
      If LenB(filterText) > 0 Then params.Add "$filter=" & stdHTTP.URIEncode(filterText)
    End If
  End If

  ODataQueryToString = JoinParams(params)
End Function

 'Convert a structured OData filter object to OData filter text.
Private Function ODataFilterToString(ByVal filter As stdJSON) As String
  Dim key As String
  Dim op As String
  Dim field As String
  Dim value As Variant
  Dim arr As stdJSON
  Dim i As Long
  Dim parts As Collection

  If filter Is Nothing Then Exit Function

  If filter.Exists("and") And IsObject(filter("and")) And TypeOf filter("and") Is stdJSON Then
    Set arr = filter("and")
    Set parts = New Collection
    For i = 1 To arr.Length
      parts.Add ODataFilterToString(arr(i))
    Next
    ODataFilterToString = JoinCollection(parts, " and ")
    Exit Function
  End If

  If filter.Exists("or") And IsObject(filter("or")) And TypeOf filter("or") Is stdJSON Then
    Set arr = filter("or")
    Set parts = New Collection
    For i = 1 To arr.Length
      parts.Add ODataFilterToString(arr(i))
    Next
    ODataFilterToString = JoinCollection(parts, " or ")
    Exit Function
  End If

  If Not filter.Exists("field") Or Not filter.Exists("operator") Then Exit Function

  field = CStr(filter("field"))
  op = LCase$(CStr(filter("operator")))
  If filter.Exists("value") Then value = filter("value")

  Select Case op
    Case "eq", "ne", "gt", "ge", "lt", "le"
      ODataFilterToString = field & " " & op & " " & ODataValueToString(value)
    Case "in"
      If IsObject(value) And TypeOf value Is stdJSON Then
        Set arr = value
        Set parts = New Collection
        For i = 1 To arr.Length
          parts.Add field & " eq " & ODataValueToString(arr(i))
        Next
        ODataFilterToString = "(" & JoinCollection(parts, " or ") & ")"
      End If
    Case "substringof"
      ODataFilterToString = "substringof(" & ODataValueToString(value) & ", " & field & ")"
    Case "startswith"
      ODataFilterToString = "startswith(" & field & ", " & ODataValueToString(value) & ")"
    Case "endswith"
      ODataFilterToString = "endswith(" & field & ", " & ODataValueToString(value) & ")"
    Case "length"
      ODataFilterToString = "length(" & field & ")"
  End Select
End Function

 'Convert a VBA value into an OData-compatible literal string.
Private Function ODataValueToString(ByVal value As Variant) As String
  If IsNull(value) Then
    ODataValueToString = "null"
  ElseIf VarType(value) = vbBoolean Then
    ODataValueToString = LCase$(CStr(value))
  ElseIf IsDate(value) Then
    ODataValueToString = FormatDateForSharePoint(value)
  ElseIf IsNumeric(value) Then
    ODataValueToString = CStr(value)
  Else
    ODataValueToString = "'" & Replace(CStr(value), "'", "''") & "'"
  End If
End Function

 'Normalize SharePoint getter payload values into client-facing shape.
Private Function TransformGetterItem(ByVal item As stdJSON) As stdJSON
  Dim out As stdJSON
  Dim idx As Long
  Dim def As TFieldDef
  Dim fieldName As String
  Dim fieldType As Long
  Dim rawValue As Variant

  Set out = JsonClone(item)

  For idx = 1 To This.FieldsCount
    def = This.Fields(idx)
    fieldName = def.InternalName
    fieldType = CLng(def.FieldType)
    If Not out.Exists(fieldName) Then GoTo ContinueField

    rawValue = out(fieldName)
    Select Case fieldType
      Case SharePointFieldMultiChoice
        If IsObject(rawValue) Then
          If TypeOf rawValue Is stdJSON Then
            If rawValue.JsonType = eJSONObject And rawValue.Exists("results") Then
              JsonAddOrSet out, fieldName, rawValue("results")
            End If
          End If
        End If
      Case SharePointFieldPerson
        NormalizePersonEmailObject rawValue
      Case SharePointFieldMultiPerson
        If IsObject(rawValue) Then
          If TypeOf rawValue Is stdJSON Then
            If rawValue.JsonType = eJSONObject And rawValue.Exists("results") Then JsonAddOrSet out, fieldName, rawValue("results")
          End If
        End If
        NormalizePersonArray out(fieldName)
    End Select
ContinueField:
  Next idx

  Set TransformGetterItem = out
End Function

 'Transform client setter payload into SharePoint REST wire format.
Private Function TransformSetterData(ByVal data As stdJSON) As stdJSON
  Dim out As stdJSON
  Dim info As Collection
  Dim node As Object
  Dim i As Long
  Dim key As String
  Dim lowerKey As String
  Dim def As TFieldDef
  Dim fieldIdx As Long
  Dim val As Variant
  Dim baseField As String
  Dim slashPos As Long
  Dim expandedProp As String

  If data Is Nothing Then
    Set TransformSetterData = stdJSON.Create(eJSONObject)
    Exit Function
  End If

  Set out = stdJSON.Create(eJSONObject)
  Set info = data.ChildrenInfo()

  For i = 1 To info.Count
    Set node = info(i)
    key = CStr(node("key"))
    lowerKey = LCase$(key)
    val = node("value")

    'Preferred expanded-field alias syntax: Owner/Email
    slashPos = InStr(1, key, "/", vbTextCompare)
    If slashPos > 1 Then
      baseField = Left$(key, slashPos - 1)
      expandedProp = LCase$(Mid$(key, slashPos + 1))
      fieldIdx = FieldIndexByInternal(baseField)
      If fieldIdx > 0 Then
        def = This.Fields(fieldIdx)
        If expandedProp = "email" Then
          If CLng(def.FieldType) = SharePointFieldPerson Then
            JsonAddOrSet out, baseField & "Id", ResolvePrincipalToUserId(val)
            GoTo ContinueNode
          ElseIf CLng(def.FieldType) = SharePointFieldMultiPerson Then
            JsonAddOrSet out, baseField & "Id", BuildIdResultsObject(val)
            GoTo ContinueNode
          End If
        ElseIf expandedProp = "id" Then
          If CLng(def.FieldType) = SharePointFieldPerson Then
            JsonAddOrSet out, baseField & "Id", ResolvePrincipalToUserId(val)
            GoTo ContinueNode
          ElseIf CLng(def.FieldType) = SharePointFieldMultiPerson Then
            JsonAddOrSet out, baseField & "Id", BuildIdResultsObject(val)
            GoTo ContinueNode
          ElseIf CLng(def.FieldType) = SharePointFieldLookup Then
            JsonAddOrSet out, baseField & "Id", NormalizeLookupId(val)
            GoTo ContinueNode
          End If
        End If
      End If
    End If

    fieldIdx = FieldIndexByInternal(lowerKey)
    If fieldIdx > 0 Then
      def = This.Fields(fieldIdx)
      Select Case CLng(def.FieldType)
        Case SharePointFieldPerson
          JsonAddOrSet out, key & "Id", ResolvePrincipalToUserId(val)
        Case SharePointFieldMultiPerson
          JsonAddOrSet out, key & "Id", BuildIdResultsObject(val)
        Case SharePointFieldMultiChoice
          JsonAddOrSet out, key, BuildStringResultsObject(val)
        Case SharePointFieldLookup
          JsonAddOrSet out, key & "Id", NormalizeLookupId(val)
        Case SharePointFieldDate
          If IsNull(val) Then
            JsonAddOrSet out, key, Null
          Else
            JsonAddOrSet out, key, FormatDateForSharePoint(val)
          End If
        Case Else
          JsonAddOrSet out, key, val
      End Select
    Else
      JsonAddOrSet out, key, val
    End If
ContinueNode:
  Next

  Set TransformSetterData = out
End Function

 'Get and cache SharePoint form digest for write/batch requests.
Private Function GetFormDigest() As String
  Dim nowUtc As Date
  Dim json As stdJSON
  Dim d As stdJSON
  Dim info As stdJSON
  Dim timeoutSec As Long

  nowUtc = Now()
  If LenB(This.FormDigestValue) > 0 And nowUtc < This.FormDigestExpiresAt Then
    GetFormDigest = This.FormDigestValue
    Exit Function
  End If

  Set json = RequestJson(This.SiteUrl & "/_api/contextinfo", "POST", vbNullString, Empty, SP_JSON_VERBOSE)
  Set d = JsonGetObject(json, "d")
  Set info = JsonGetObject(d, "GetContextWebInformation")
  This.FormDigestValue = CStr(info("FormDigestValue"))

  timeoutSec = CLng(info("FormDigestTimeoutSeconds"))
  If timeoutSec < DIGEST_EXPIRY_BUFFER_SECONDS + 1 Then timeoutSec = 1800
  This.FormDigestExpiresAt = DateAdd("s", timeoutSec - DIGEST_EXPIRY_BUFFER_SECONDS, nowUtc)

  GetFormDigest = This.FormDigestValue
End Function

 'Get and cache SharePoint list item entity type.
Private Function GetListItemEntityType() As String
  Dim json As stdJSON
  Dim d As stdJSON

  If LenB(This.ListItemEntityType) > 0 Then
    GetListItemEntityType = This.ListItemEntityType
    Exit Function
  End If

  Set json = RequestJson(ListEndpoint("?$select=ListItemEntityTypeFullName"), "GET")
  Set d = JsonGetObject(json, "d")
  If d Is Nothing Or Not d.Exists("ListItemEntityTypeFullName") Then
    Err.Raise 5, "SharepointList::GetListItemEntityType", "ListItemEntityTypeFullName was not found."
  End If
  This.ListItemEntityType = CStr(d("ListItemEntityTypeFullName"))
  GetListItemEntityType = This.ListItemEntityType
End Function

 'Attach `__metadata.type` required for list write operations.
Private Sub SetMetadataType(ByVal data As stdJSON, ByVal itemType As String)
  Dim m As stdJSON
  Set m = stdJSON.Create(eJSONObject)
  m.Add "type", itemType
  JsonAddOrSet data, "__metadata", m
End Sub

 'Parse multipart OData batch response into status/json result rows.
Private Function ParseODataBatchResponse(ByVal body As String) As Collection
  Dim parts() As String
  Dim boundary As String
  Dim i As Long
  Dim part As String
  Dim matches As Object
  Dim re As Object
  Dim row As Object
  Dim statusCode As Long
  Dim statusText As String

  Set ParseODataBatchResponse = New Collection
  boundary = ExtractBatchBoundary(body)
  If LenB(boundary) = 0 Then Err.Raise 5, "SharepointList::ParseODataBatchResponse", "No batch boundary found."

  parts = Split(body, boundary)
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "HTTP/\d\.\d\s+(\d+)\s+([^\r\n]+)"
  re.Global = False
  re.IgnoreCase = True

  For i = LBound(parts) To UBound(parts)
    part = Trim$(parts(i))
    If LenB(part) = 0 Then GoTo ContinuePart
    If part = "--" Then GoTo ContinuePart

    If re.Test(part) Then
      Set matches = re.Execute(part)
      statusCode = CLng(matches(0).SubMatches(0))
      statusText = CStr(matches(0).SubMatches(1))
    Else
      statusCode = 0
      statusText = vbNullString
    End If

    Set row = NewDictionary(vbTextCompare)
    row("statusCode") = statusCode
    row("statusText") = statusText
    set row("json") = TryExtractJsonObject(part)
    ParseODataBatchResponse.Add row
ContinuePart:
  Next
End Function

 'Try to parse the first JSON object found in a batch response part.
Private Function TryExtractJsonObject(ByVal bodyPart As String) As stdJSON
  Dim pStart As Long
  Dim pEnd As Long
  Dim jsonText As String

  pStart = InStr(1, bodyPart, "{")
  pEnd = InStrRev(bodyPart, "}")
  If pStart <= 0 Or pEnd <= pStart Then Exit Function

  jsonText = Mid$(bodyPart, pStart, pEnd - pStart + 1)
  On Error Resume Next
  Set TryExtractJsonObject = stdJSON.CreateFromString(jsonText)
  On Error GoTo 0
End Function

 'Extract the batch response boundary marker from raw response text.
Private Function ExtractBatchBoundary(ByVal body As String) As String
  Dim re As Object
  Dim m As Object

  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "--batchresponse_[\w\-]+"
  re.Global = False
  re.IgnoreCase = True
  If re.Test(body) Then
    Set m = re.Execute(body)(0)
    ExtractBatchBoundary = m.Value
  End If
End Function

 'Split a JSON array into chunked JSON arrays of a given size.
Private Function SplitJsonArrayIntoChunks(ByVal inputArr As stdJSON, ByVal chunkSize As Long) As Collection
  Dim out As Collection
  Dim i As Long
  Dim chunk As stdJSON

  Set out = New Collection
  If inputArr Is Nothing Then
    Set SplitJsonArrayIntoChunks = out
    Exit Function
  End If

  For i = 1 To inputArr.Length
    If (i - 1) Mod chunkSize = 0 Then
      Set chunk = stdJSON.Create(eJSONArray)
      out.Add chunk
    End If
    chunk.Add inputArr(i)
  Next

  Set SplitJsonArrayIntoChunks = out
End Function

 'Compare values with field-aware normalization (for simplification).
Private Function ValuesEqualForField(ByVal key As String, ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
  Dim fieldType As Long
  fieldType = ResolveFieldTypeByName(key)

  If fieldType = SharePointFieldDate Then
    ValuesEqualForField = (NormalizeDateForCompare(leftValue) = NormalizeDateForCompare(rightValue))
  ElseIf IsObject(leftValue) Xor IsObject(rightValue) Then
    ValuesEqualForField = False
  ElseIf IsObject(leftValue) And IsObject(rightValue) Then
    If TypeOf leftValue Is stdJSON And TypeOf rightValue Is stdJSON Then
      ValuesEqualForField = (leftValue.ToString() = rightValue.ToString())
    Else
      ValuesEqualForField = (TypeName(leftValue) = TypeName(rightValue))
    End If
  ElseIf IsNull(leftValue) Xor IsNull(rightValue) Then
    ValuesEqualForField = False
  Else
    ValuesEqualForField = (CStr(leftValue) = CStr(rightValue))
  End If
End Function

 'Normalize date-like values for robust comparison operations.
Private Function NormalizeDateForCompare(ByVal value As Variant) As Double
  On Error GoTo invalidDate
  If IsDate(value) Then
    NormalizeDateForCompare = CDbl(CDate(value))
    Exit Function
  End If
  If VarType(value) = vbString Then
    If InStr(1, value, "T", vbTextCompare) > 0 Then
      NormalizeDateForCompare = CDbl(CDate(Replace(Replace(CStr(value), "T", " "), "Z", vbNullString)))
      Exit Function
    End If
    NormalizeDateForCompare = CDbl(CDate(value))
    Exit Function
  End If
invalidDate:
  NormalizeDateForCompare = 0#
End Function

 'Normalize lookup input into a lookup item id value.
Private Function NormalizeLookupId(ByVal value As Variant) As Variant
  If IsObject(value) Then
    If TypeOf value Is stdJSON Then
      If value.Exists("Id") Then
        NormalizeLookupId = CLng(value("Id"))
        Exit Function
      End If
    End If
  End If
  NormalizeLookupId = value
End Function

 'Build SharePoint multichoice payload `{__metadata, results}`.
Private Function BuildStringResultsObject(ByVal value As Variant) As stdJSON
  Dim out As stdJSON
  Dim meta As stdJSON
  Dim results As stdJSON
  Dim i As Long

  Set out = stdJSON.Create(eJSONObject)
  Set meta = stdJSON.Create(eJSONObject)
  meta.Add "type", "Collection(Edm.String)"
  out.Add "__metadata", meta

  Set results = stdJSON.Create(eJSONArray)
  If IsObject(value) And TypeOf value Is stdJSON Then
    If value.JsonType = eJSONArray Then
      For i = 1 To value.Length
        results.Add CStr(value(i))
      Next
    ElseIf value.JsonType = eJSONObject And value.Exists("results") Then
      If IsObject(value("results")) And TypeOf value("results") Is stdJSON Then
        For i = 1 To value("results").Length
          results.Add CStr(value("results")(i))
        Next
      End If
    End If
  ElseIf IsArray(value) Then
    For i = LBound(value) To UBound(value)
      results.Add CStr(value(i))
    Next
  ElseIf LenB(CStr(value)) > 0 Then
    results.Add CStr(value)
  End If

  out.Add "results", results
  Set BuildStringResultsObject = out
End Function

 'Build SharePoint person/multiperson id payload `{results:[ids...]}`.
Private Function BuildIdResultsObject(ByVal value As Variant) As stdJSON
  Dim out As stdJSON
  Dim results As stdJSON
  Dim i As Long

  Set out = stdJSON.Create(eJSONObject)
  Set results = stdJSON.Create(eJSONArray)

  If IsObject(value) And TypeOf value Is stdJSON Then
    If value.JsonType = eJSONArray Then
      For i = 1 To value.Length
        results.Add ResolvePrincipalToUserId(value(i))
      Next
    ElseIf value.JsonType = eJSONObject And value.Exists("results") Then
      If IsObject(value("results")) And TypeOf value("results") Is stdJSON Then
        For i = 1 To value("results").Length
          results.Add ResolvePrincipalToUserId(value("results")(i))
        Next
      End If
    End If
  ElseIf IsArray(value) Then
    For i = LBound(value) To UBound(value)
      results.Add ResolvePrincipalToUserId(value(i))
    Next
  ElseIf Not IsNull(value) And LenB(CStr(value)) > 0 Then
    results.Add ResolvePrincipalToUserId(value)
  End If

  out.Add "results", results
  Set BuildIdResultsObject = out
End Function

 'Resolve principal input (id/object/email) to a SharePoint user id.
Private Function ResolvePrincipalToUserId(ByVal value As Variant) As Long
  If IsNumeric(value) Then
    ResolvePrincipalToUserId = CLng(value)
    Exit Function
  End If
  If IsObject(value) Then
    If TypeOf value Is stdJSON Then
      If value.Exists("Id") And IsNumeric(value("Id")) Then
        ResolvePrincipalToUserId = CLng(value("Id"))
        Exit Function
      End If
      If value.Exists("Email") Then
        ResolvePrincipalToUserId = GetOrCreateUserId(CStr(value("Email")))
        Exit Function
      End If
      If value.Exists("EMail") Then
        ResolvePrincipalToUserId = GetOrCreateUserId(CStr(value("EMail")))
        Exit Function
      End If
    End If
  End If
  ResolvePrincipalToUserId = GetOrCreateUserId(CStr(value))
End Function

 'Format a date value into SharePoint datetime string format.
Private Function FormatDateForSharePoint(ByVal value As Variant) As String
  Dim dt As Date
  dt = CDate(value)
  FormatDateForSharePoint = Format$(dt, "yyyy-mm-dd hh:nn:ss")
End Function

 'Normalize person object key naming (`EMail` -> `Email`).
Private Sub NormalizePersonEmailObject(ByRef value As Variant)
  If Not IsObject(value) Then Exit Sub
  If Not TypeOf value Is stdJSON Then Exit Sub
  If value.Exists("EMail") And Not value.Exists("Email") Then value.Add "Email", value("EMail")
End Sub

 'Apply person key normalization across person arrays.
Private Sub NormalizePersonArray(ByVal value As Variant)
  Dim i As Long
  Dim p As Variant
  If Not IsObject(value) Then Exit Sub
  If Not TypeOf value Is stdJSON Then Exit Sub
  If value.JsonType <> eJSONArray Then Exit Sub
  For i = 1 To value.Length
    p = value(i)
    NormalizePersonEmailObject p
  Next
End Sub

'Find schema index by internal field name.
'@returns - 1-based index in `This.Fields`, or 0 when not found.
Private Function FieldIndexByInternal(ByVal internalName As String) As Long
  Dim target As String
  Dim i As Long

  target = LCase$(Trim$(internalName))
  If LenB(target) = 0 Then Exit Function

  For i = 1 To This.FieldsCount
    If LCase$(This.Fields(i).InternalName) = target Then
      FieldIndexByInternal = i
      Exit Function
    End If
  Next
End Function

 'Resolve field type by exact/expanded/derived field naming forms.
Private Function ResolveFieldTypeByName(ByVal key As String) As Long
  Dim lower As String
  Dim slashPos As Long
  Dim candidate As String
  Dim fieldIdx As Long
  lower = LCase$(key)

  fieldIdx = FieldIndexByInternal(lower)
  If fieldIdx > 0 Then
    ResolveFieldTypeByName = CLng(This.Fields(fieldIdx).FieldType)
    Exit Function
  End If

  slashPos = InStr(1, lower, "/", vbTextCompare)
  If slashPos > 1 Then
    candidate = Left$(lower, slashPos - 1)
    fieldIdx = FieldIndexByInternal(candidate)
    If fieldIdx > 0 Then
      ResolveFieldTypeByName = CLng(This.Fields(fieldIdx).FieldType)
      Exit Function
    End If
  End If

  If Right$(lower, 2) = "id" Then
    candidate = Left$(lower, Len(lower) - 2)
    fieldIdx = FieldIndexByInternal(candidate)
    If fieldIdx > 0 Then
      ResolveFieldTypeByName = CLng(This.Fields(fieldIdx).FieldType)
      Exit Function
    End If
  End If

  ResolveFieldTypeByName = SharePointFieldText
End Function

 'Build default schema-driven `$select` and `$expand` fragments.
Private Sub BuildDefaultSelectExpand(ByRef selectCsv As String, ByRef expandCsv As String)
  Dim selectParts As Collection
  Dim expandParts As Collection
  Dim fieldDef As TFieldDef
  Dim internalName As String
  Dim fType As Long
  Dim props As Collection
  Dim i As Long
  Dim idx As Long

  Set selectParts = New Collection
  Set expandParts = New Collection
  AddUnique selectParts, stdHTTP.URIEncode("Id")

  For idx = 1 To This.FieldsCount
    fieldDef = This.Fields(idx)
    internalName = fieldDef.InternalName
    fType = CLng(fieldDef.FieldType)

    Select Case fType
      Case SharePointFieldPerson, SharePointFieldMultiPerson
        AddUnique expandParts, stdHTTP.URIEncode(internalName)
        Set props = fieldDef.SelectProps
        If props.Count = 0 Then
          AddUnique selectParts, stdHTTP.URIEncode(internalName & "/Id")
          AddUnique selectParts, stdHTTP.URIEncode(internalName & "/Title")
          AddUnique selectParts, stdHTTP.URIEncode(internalName & "/EMail")
        Else
          For i = 1 To props.Count
            AddUnique selectParts, stdHTTP.URIEncode(internalName & "/" & CStr(props(i)))
          Next
        End If
      Case SharePointFieldLookup, SharePointFieldFile
        AddUnique expandParts, stdHTTP.URIEncode(internalName)
        Set props = fieldDef.SelectProps
        If props.Count = 0 Then
          AddUnique selectParts, stdHTTP.URIEncode(internalName & "/Id")
          AddUnique selectParts, stdHTTP.URIEncode(internalName & "/Title")
        Else
          For i = 1 To props.Count
            AddUnique selectParts, stdHTTP.URIEncode(internalName & "/" & CStr(props(i)))
          Next
        End If
      Case Else
        AddUnique selectParts, stdHTTP.URIEncode(internalName)
    End Select
  Next idx

  selectCsv = JoinCollection(selectParts, ",")
  expandCsv = JoinCollection(expandParts, ",")
End Sub

 'Encode a JSON array of identifiers as a comma-separated value list.
Private Function EncodeArrayCsv(ByVal arr As stdJSON) As String
  Dim items As Collection
  Dim i As Long
  Set items = New Collection
  For i = 1 To arr.Length
    AddUnique items, stdHTTP.URIEncode(CStr(arr(i)))
  Next
  EncodeArrayCsv = JoinCollection(items, ",")
End Function

 'Build OData `$orderby` text from order descriptor objects.
Private Function BuildOrderByString(ByVal arr As stdJSON) As String
  Dim parts As Collection
  Dim i As Long
  Dim obj As stdJSON
  Dim segment As String

  Set parts = New Collection
  For i = 1 To arr.Length
    If IsObject(arr(i)) And TypeOf arr(i) Is stdJSON Then
      Set obj = arr(i)
      If obj.Exists("name") Then
        segment = stdHTTP.URIEncode(CStr(obj("name")))
        If obj.Exists("direction") Then segment = segment & " " & LCase$(CStr(obj("direction")))
        parts.Add segment
      End If
    End If
  Next
  BuildOrderByString = JoinCollection(parts, ", ")
End Function

 'Join query parameters collection into a `?k=v&...` querystring.
Private Function JoinParams(ByVal params As Collection) As String
  If params.Count = 0 Then Exit Function
  JoinParams = "?" & JoinCollection(params, "&")
End Function

 'Normalize raw query text to either empty or `?`-prefixed form.
Private Function NormalizeRawQuery(ByVal raw As String) As String
  raw = Trim$(raw)
  If LenB(raw) = 0 Then Exit Function
  If Left$(raw, 1) = "?" Then
    NormalizeRawQuery = raw
  Else
    NormalizeRawQuery = "?" & raw
  End If
End Function

 'Join collection items into a delimited string.
Private Function JoinCollection(ByVal items As Collection, ByVal delim As String) As String
  Dim i As Long
  Dim sb As String
  For i = 1 To items.Count
    If i > 1 Then sb = sb & delim
    sb = sb & CStr(items(i))
  Next
  JoinCollection = sb
End Function

 'Add a string value to collection only if not already present.
Private Sub AddUnique(ByVal list As Collection, ByVal value As String)
  Dim i As Long
  For i = 1 To list.Count
    If CStr(list(i)) = value Then Exit Sub
  Next
  list.Add value
End Sub

 'Parse list page URL into site URL, list title, and relative path.
Private Sub ParseListUrl(ByVal raw As String, ByRef siteUrl As String, ByRef listTitle As String, ByRef listRelativePath As String)
  Dim cleaned As String
  Dim listsPos As Long
  Dim afterLists As String
  Dim slashPos As Long
  Dim listPathSegment As String

  cleaned = Trim$(raw)
  cleaned = Split(cleaned, "?")(0)
  cleaned = Split(cleaned, "#")(0)
  If Right$(cleaned, 1) = "/" Then cleaned = Left$(cleaned, Len(cleaned) - 1)

  listsPos = InStr(1, LCase$(cleaned), "/lists/", vbTextCompare)
  If listsPos <= 0 Then Err.Raise 5, "SharepointList::ParseListUrl", "ListUrl must contain `/Lists/<ListName>`."

  siteUrl = Left$(cleaned, listsPos - 1)
  afterLists = Mid$(cleaned, listsPos + Len("/Lists/"))
  slashPos = InStr(1, afterLists, "/", vbTextCompare)
  If slashPos > 0 Then
    listPathSegment = Left$(afterLists, slashPos - 1)
  Else
    listPathSegment = afterLists
  End If

  listTitle = stdHTTP.URIDecode(listPathSegment)
  siteUrl = TrimTrailingSlash(siteUrl)
  listRelativePath = BuildServerRelativeListPath(siteUrl, listPathSegment)
End Sub

Private Sub EnsureListDisplayTitle()
  Dim resolvedTitle As String

  If This.ListTitleResolved Then Exit Sub

  resolvedTitle = ResolveListDisplayTitle()
  If LenB(Trim$(resolvedTitle)) = 0 Then
    Err.Raise 5, "SharepointList::EnsureListDisplayTitle", "Unable to resolve list display title from list URL."
  End If

  This.ListTitle = resolvedTitle
  This.ListTitleResolved = True
End Sub

Private Function ResolveListDisplayTitle() As String
  Dim api As String
  Dim json As stdJSON
  Dim d As stdJSON

  api = This.SiteUrl & "/_api/web/GetList('" & EscapeODataLiteral(This.ListRelativePath) & "')?$select=Title"
  Set json = RequestJson(api, "GET")
  Set d = JsonGetObject(json, "d")

  If d Is Nothing Or Not d.Exists("Title") Then
    Err.Raise 5, "SharepointList::ResolveListDisplayTitle", "SharePoint did not return a list Title."
  End If

  ResolveListDisplayTitle = CStr(d("Title"))
End Function

Private Function BuildServerRelativeListPath(ByVal siteUrl As String, ByVal listPathSegment As String) As String
  Dim siteRelativePath As String
  Dim decodedSegment As String

  siteRelativePath = SiteServerRelativePath(siteUrl)
  decodedSegment = stdHTTP.URIDecode(listPathSegment)

  If siteRelativePath = "/" Then
    BuildServerRelativeListPath = "/Lists/" & decodedSegment
  Else
    BuildServerRelativeListPath = siteRelativePath & "/Lists/" & decodedSegment
  End If
End Function

Private Function SiteServerRelativePath(ByVal absoluteSiteUrl As String) As String
  Dim stripped As String
  Dim schemePos As Long
  Dim pathPos As Long

  stripped = Trim$(absoluteSiteUrl)
  schemePos = InStr(1, stripped, "://", vbTextCompare)
  If schemePos > 0 Then stripped = Mid$(stripped, schemePos + 3)

  pathPos = InStr(1, stripped, "/", vbTextCompare)
  If pathPos <= 0 Then
    SiteServerRelativePath = "/"
    Exit Function
  End If

  SiteServerRelativePath = Mid$(stripped, pathPos)
  If LenB(SiteServerRelativePath) = 0 Then
    SiteServerRelativePath = "/"
  ElseIf Right$(SiteServerRelativePath, 1) = "/" And Len(SiteServerRelativePath) > 1 Then
    SiteServerRelativePath = Left$(SiteServerRelativePath, Len(SiteServerRelativePath) - 1)
  End If
End Function

 'Trim trailing slash characters from a URL/string.
Private Function TrimTrailingSlash(ByVal s As String) As String
  Do While Right$(s, 1) = "/"
    s = Left$(s, Len(s) - 1)
  Loop
  TrimTrailingSlash = s
End Function

 'Escape single quotes for OData string literal segments.
Private Function EscapeODataLiteral(ByVal s As String) As String
  EscapeODataLiteral = Replace(s, "'", "''")
End Function

 'Encode a plain string as a JSON string literal.
Private Function JsonStringLiteral(ByVal s As String) As String
  s = Replace(s, "\", "\\")
  s = Replace(s, """", "\""")
  JsonStringLiteral = """" & s & """"
End Function

 'Get a child object by key, returning Nothing when absent/non-object.
Private Function JsonGetObject(ByVal parent As stdJSON, ByVal key As String) As stdJSON
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then
    If TypeOf parent(key) Is stdJSON Then Set JsonGetObject = parent(key)
  End If
End Function

 'Add or set a key/value on a JSON object.
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

 'Create a deep JSON clone using `ToVBObject` round-trip.
Private Function JsonClone(ByVal source As stdJSON) As stdJSON
  If source Is Nothing Then
    Set JsonClone = stdJSON.Create(eJSONObject)
  Else
    Set JsonClone = stdJSON.CreateFromVariant(source.ToVBObject())
  End If
End Function

 'Normalize select-props input into a collection of strings.
Private Function NormalizeSelectProps(ByVal SelectProps As Variant) As Collection
  Dim out As Collection
  Dim i As Long
  Dim arr() As String

  Set out = New Collection
  If IsMissingOrEmpty(SelectProps) Then
    Set NormalizeSelectProps = out
    Exit Function
  End If

  If IsObject(SelectProps) Then
    If TypeOf SelectProps Is stdJSON Then
      If SelectProps.JsonType = eJSONArray Then
        For i = 1 To SelectProps.Length
          out.Add CStr(SelectProps(i))
        Next
      End If
      Set NormalizeSelectProps = out
      Exit Function
    ElseIf TypeName(SelectProps) = "Collection" Then
      For i = 1 To SelectProps.Count
        out.Add CStr(SelectProps(i))
      Next
      Set NormalizeSelectProps = out
      Exit Function
    End If
  End If

  If IsArray(SelectProps) Then
    For i = LBound(SelectProps) To UBound(SelectProps)
      out.Add CStr(SelectProps(i))
    Next
  ElseIf VarType(SelectProps) = vbString Then
    arr = Split(CStr(SelectProps), ",")
    For i = LBound(arr) To UBound(arr)
      If LenB(Trim$(arr(i))) > 0 Then out.Add Trim$(arr(i))
    Next
  End If

  Set NormalizeSelectProps = out
End Function

'Check whether an optional variant is missing/empty/nothing.
Private Function IsMissingOrEmpty(Optional ByVal v As Variant) As Boolean
  If IsMissing(v) Then
    IsMissingOrEmpty = True
    Exit Function
  End If

  If IsObject(v) Then
    IsMissingOrEmpty = (v Is Nothing)
  Else
    IsMissingOrEmpty = IsEmpty(v)
  End If
End Function

 'Create a late-bound dictionary with optional compare mode.
Private Function NewDictionary(Optional ByVal compareMode As Long = vbBinaryCompare) As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
  On Error Resume Next
  NewDictionary.CompareMode = compareMode
  On Error GoTo 0
End Function

 'Generate a GUID string suitable for batch/change set ids.
Private Function GenerateUuid() As String
  Static seeded As Boolean
  Dim template As String
  Dim i As Long
  Dim ch As String
  Dim out As String

  If Not seeded Then
    Randomize
    seeded = True
  End If

  template = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  For i = 1 To Len(template)
    ch = Mid$(template, i, 1)
    Select Case ch
      Case "x"
        out = out & Hex$(Int(Rnd() * 16))
      Case "y"
        out = out & Hex$(8 + Int(Rnd() * 4))
      Case Else
        out = out & ch
    End Select
  Next

  GenerateUuid = LCase$(out)
End Function

 'Attempt to discover list fields from SharePoint metadata.
Private Sub TryAutoDiscoverFields()
  Dim api As String
  Dim json As stdJSON
  Dim d As stdJSON
  Dim rows As stdJSON
  Dim i As Long
  Dim row As stdJSON
  Dim fType As SharePointFieldType
  Dim typeName As String

  api = ListEndpoint("/fields?$select=Title,InternalName,TypeAsString,AllowMultipleValues")
  Set json = RequestJson(api, "GET")
  Set d = JsonGetObject(json, "d")
  If d Is Nothing Then Exit Sub
  If Not d.Exists("results") Then Exit Sub
  If Not IsObject(d("results")) Then Exit Sub
  If Not TypeOf d("results") Is stdJSON Then Exit Sub
  Set rows = d("results")

  For i = 1 To rows.Length
    If Not IsObject(rows(i)) Then GoTo ContinueRow
    If Not TypeOf rows(i) Is stdJSON Then GoTo ContinueRow
    Set row = rows(i)
    If Not row.Exists("InternalName") Then GoTo ContinueRow

    typeName = LCase$(CStr(row("TypeAsString")))
    Select Case typeName
      Case "text", "note": fType = SharePointFieldText
      Case "number", "currency", "integer", "counter": fType = SharePointFieldNumber
      Case "boolean": fType = SharePointFieldBoolean
      Case "datetime": fType = SharePointFieldDate
      Case "choice": fType = SharePointFieldChoice
      Case "multichoice": fType = SharePointFieldMultiChoice
      Case "user"
        If row.Exists("AllowMultipleValues") And CBool(row("AllowMultipleValues")) Then
          fType = SharePointFieldMultiPerson
        Else
          fType = SharePointFieldPerson
        End If
      Case "lookup"
        If row.Exists("AllowMultipleValues") And CBool(row("AllowMultipleValues")) Then
          fType = SharePointFieldLookup
        Else
          fType = SharePointFieldLookup
        End If
      Case Else
        fType = SharePointFieldText
    End Select

    Call FieldsAdd(CStr(row("InternalName")), fType)
ContinueRow:
  Next
End Sub
