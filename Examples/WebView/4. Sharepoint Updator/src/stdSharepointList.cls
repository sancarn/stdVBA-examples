VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSharepointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A SharePoint list REST client built on `stdHTTP`, `stdJSON`, and `SharepointAuthenticator`.
'It supports schema-aware field transforms (person/multiperson/multichoice/lookup/date),
'single-item and paged reads, in-place search, item create/update, batch operations, and
'commit-payload orchestration with optional payload simplification.
'@aidate 2026-11-26
'@aidesc AI added batch response parsing including person lookups, and timed-out update verification.
'@example ```vb
'Dim auth As SharepointAuthenticator
'Set auth = SharepointAuthenticator.Create("https://contoso.sharepoint.com")
'
'Dim risks As stdSharepointList
'Set risks = stdSharepointList.Create( _
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

#If Mac Then
  Private Declare PtrSafe Sub uuid_generate_random Lib "/usr/lib/libSystem.B.dylib" (ByRef firstByte As Byte)
#ElseIf VBA7 Then
  Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef firstByte As Byte) As Long
#Else
  Private Declare Function CoCreateGuid Lib "ole32" (ByRef firstByte As Byte) As Long
#End If

Private Type TFieldDef
  InternalName As String
  FieldType As SharePointFieldType
  Nullable As Boolean
  LookupField As String
  SelectProps As Collection
End Type

Private Type TThis
  SiteURL As String
  ListTitle As String
  listRelativePath As String
  ListTitleResolved As Boolean

  Authenticator As stdICallable
  fields() As TFieldDef
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
Private Const PEOPLE_PICKER_MAX_RESULTS_CAP As Long = 200

'Create a SharePoint list client from a SharePoint list URL.
'@constructor
'@param ListUrl - List page URL (for example `/sites/X/Lists/MyList/AllItems.aspx`).
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Variant)=>Void> - Authenticator used by `stdHTTP`.
'@param AutoDiscoverFields - If true, attempts to infer fields from SharePoint metadata.
'@returns - A configured `stdSharepointList` instance.
Public Function Create( _
  ByVal ListUrl As String, _
  ByVal Authenticator As stdICallable, _
  Optional ByVal AutoDiscoverFields As Boolean = False _
) As stdSharepointList
  Dim instance As stdSharepointList
  Set instance = New stdSharepointList
  Call instance.protInit(ListUrl, Authenticator, AutoDiscoverFields)
  Set Create = instance
End Function

'Create a SharePoint list client from a site URL and list display title.
'@constructor
'@param SiteUrl - Absolute SharePoint site URL (for example `https://contoso.sharepoint.com/sites/Projects`).
'@param ListTitle - List display title used by `GetByTitle`.
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Variant)=>Void> - Authenticator used by `stdHTTP`.
'@param AutoDiscoverFields - If true, attempts to infer fields from SharePoint metadata.
'@returns - A configured `stdSharepointList` instance.
Public Function CreateFromTitle( _
  ByVal SiteURL As String, _
  ByVal ListTitle As String, _
  ByVal Authenticator As stdICallable, _
  Optional ByVal AutoDiscoverFields As Boolean = False _
) As stdSharepointList
  Dim instance As stdSharepointList
  Set instance = New stdSharepointList
  Call instance.protInitFromTitle(SiteURL, ListTitle, Authenticator, AutoDiscoverFields)
  Set CreateFromTitle = instance
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
  Dim SiteURL As String
  Dim ListTitle As String
  Dim relPath As String

  If LenB(Trim$(ListUrl)) = 0 Then
    Err.Raise 5, "stdSharepointList::protInit", "ListUrl cannot be blank."
  End If
  If Authenticator Is Nothing Then
    Err.Raise 5, "stdSharepointList::protInit", "Authenticator cannot be Nothing."
  End If

  Call ParseListUrl(ListUrl, SiteURL, ListTitle, relPath)

  This.SiteURL = SiteURL
  This.ListTitle = ListTitle
  This.listRelativePath = relPath
  This.ListTitleResolved = False
  Set This.Authenticator = Authenticator
  This.FieldsCount = 0
  Erase This.fields
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

'Initialize list state from an absolute site URL and list display title.
'@protected
'@param SiteUrl - Absolute SharePoint site URL.
'@param ListTitle - List display title used by `GetByTitle`.
'@param Authenticator - Authenticator used by outgoing HTTP requests.
'@param AutoDiscoverFields - If true, attempts to infer fields from SharePoint metadata.
Public Sub protInitFromTitle( _
  ByVal SiteURL As String, _
  ByVal ListTitle As String, _
  ByVal Authenticator As stdICallable, _
  Optional ByVal AutoDiscoverFields As Boolean = False _
)
  Dim cleanedSite As String
  Dim cleanedTitle As String

  cleanedSite = TrimTrailingSlash(Trim$(SiteURL))
  cleanedTitle = Trim$(ListTitle)

  If LenB(cleanedSite) = 0 Then
    Err.Raise 5, "stdSharepointList::protInitFromTitle", "SiteUrl cannot be blank."
  End If
  If LenB(cleanedTitle) = 0 Then
    Err.Raise 5, "stdSharepointList::protInitFromTitle", "ListTitle cannot be blank."
  End If
  If Authenticator Is Nothing Then
    Err.Raise 5, "stdSharepointList::protInitFromTitle", "Authenticator cannot be Nothing."
  End If

  This.SiteURL = cleanedSite
  This.ListTitle = cleanedTitle
  This.listRelativePath = vbNullString
  This.ListTitleResolved = True
  Set This.Authenticator = Authenticator
  This.FieldsCount = 0
  Erase This.fields
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
Public Function FieldsClear() As stdSharepointList
  This.FieldsCount = 0
  Erase This.fields
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
) As stdSharepointList
  Dim key As String
  Dim fieldDef As TFieldDef
  Dim idx As Long

  key = LCase$(Trim$(InternalName))
  If LenB(key) = 0 Then
    Err.Raise 5, "stdSharepointList::FieldsAdd", "InternalName cannot be blank."
  End If

  fieldDef.InternalName = Trim$(InternalName)
  fieldDef.FieldType = FieldType
  fieldDef.Nullable = CBool(Nullable)
  fieldDef.LookupField = LookupField
  Set fieldDef.SelectProps = NormalizeSelectProps(SelectProps)

  idx = FieldIndexByInternal(key)
  If idx = 0 Then
    This.FieldsCount = This.FieldsCount + 1
    ReDim Preserve This.fields(1 To This.FieldsCount)
    idx = This.FieldsCount
  End If
  This.fields(idx) = fieldDef

  Set FieldsAdd = Me
End Function

'Fetch list field metadata from SharePoint for validation and mapping.
'@returns as stdJSON<Array<{Title:string, InternalName:string, TypeAsString:string, Required:boolean, ReadOnlyField:boolean, Hidden:boolean, MaxLength:Variant, Choices:stdJSON[], AllowMultipleValues:boolean}>> - Field schema rows.
Public Function FieldsFetchSchema() As stdJSON
  Dim api As String
  Dim json As stdJSON
  Dim d As stdJSON
  Dim rows As stdJSON
  Dim outArr As stdJSON
  Dim i As Long
  Dim row As stdJSON
  Dim one As stdJSON
  Dim choices As stdJSON
  Dim choiceResults As stdJSON
  Dim j As Long

  api = ListEndpoint( _
    "/fields?$select=Title,InternalName,EntityPropertyName,TypeAsString,Required,ReadOnlyField,Hidden,MaxLength,Choices,AllowMultipleValues" _
  )
  Set json = RequestJson(api, "GET")
  Set d = JsonGetObject(json, "d")
  Set outArr = stdJSON.Create(eJSONArray)

  If d Is Nothing Then
    Set FieldsFetchSchema = outArr
    Exit Function
  End If
  If Not d.Exists("results") Then
    Set FieldsFetchSchema = outArr
    Exit Function
  End If
  If Not IsObject(d("results")) Then
    Set FieldsFetchSchema = outArr
    Exit Function
  End If
  If Not TypeOf d("results") Is stdJSON Then
    Set FieldsFetchSchema = outArr
    Exit Function
  End If

  Set rows = d("results")
  For i = 1 To rows.Length
    If Not IsObject(rows(i)) Then GoTo ContinueSchemaRow
    If Not TypeOf rows(i) Is stdJSON Then GoTo ContinueSchemaRow
    Set row = rows(i)
    If Not row.Exists("InternalName") Then GoTo ContinueSchemaRow

    Set one = stdJSON.Create(eJSONObject)
    one.Add "Title", JsonGetString(row, "Title")
    one.Add "InternalName", CStr(row("InternalName"))
    one.Add "EntityPropertyName", JsonGetString(row, "EntityPropertyName")
    one.Add "TypeAsString", JsonGetString(row, "TypeAsString")
    one.Add "Required", JsonGetBoolean(row, "Required")
    one.Add "ReadOnlyField", JsonGetBoolean(row, "ReadOnlyField")
    one.Add "Hidden", JsonGetBoolean(row, "Hidden")
    If row.Exists("MaxLength") And Not IsNull(row("MaxLength")) And IsNumeric(row("MaxLength")) Then
      one.Add "MaxLength", CLng(row("MaxLength"))
    Else
      one.Add "MaxLength", Null
    End If
    one.Add "AllowMultipleValues", JsonGetBoolean(row, "AllowMultipleValues")

    Set choices = stdJSON.Create(eJSONArray)
    If row.Exists("Choices") Then
      If IsObject(row("Choices")) And TypeOf row("Choices") Is stdJSON Then
        If row("Choices").JsonType = eJSONObject And row("Choices").Exists("results") Then
          If IsObject(row("Choices")("results")) And TypeOf row("Choices")("results") Is stdJSON Then
            Set choiceResults = row("Choices")("results")
            For j = 1 To choiceResults.Length
              choices.Add CStr(choiceResults(j))
            Next
          End If
        ElseIf row("Choices").JsonType = eJSONArray Then
          For j = 1 To row("Choices").Length
            choices.Add CStr(row("Choices")(j))
          Next
        End If
      End If
    End If
    one.Add "Choices", choices
    outArr.Add one
ContinueSchemaRow:
  Next

  Set FieldsFetchSchema = outArr
End Function

'Create an equality OData filter object.
'@param FieldName - Field/internal name.
'@param Value - Value to compare.
'@returns as stdJSON<{field: string, operator: "eq", value: Variant}> - OData filter object.
Public Function ODataFilterEq(ByVal fieldName As String, ByVal value As Variant) As stdJSON
  Set ODataFilterEq = stdJSON.Create(eJSONObject)
  Call ODataFilterEq.Add("field", fieldName)
  Call ODataFilterEq.Add("operator", "eq")
  Call ODataFilterEq.Add("value", value)
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
Public Function ODataOrderBy(ByVal fieldName As String, Optional ByVal Desc As Boolean = False) As stdJSON
  Set ODataOrderBy = stdJSON.Create(eJSONObject)
  Call ODataOrderBy.Add("name", fieldName)
  Call ODataOrderBy.Add("direction", IIf(Desc, "desc", "asc"))
End Function

'Get one list item by id.
'@param ItemId - SharePoint item id.
'@param Query as Variant<String|stdJSON> - Optional OData query (raw or structured).
'@returns as stdJSON - Item data transformed to client shape.
Public Function ItemGet(ByVal itemId As Variant, Optional ByVal query As Variant) As stdJSON
  Dim api As String
  Dim data As stdJSON
  Dim d As stdJSON

  api = ListItemsEndpoint("(" & CStr(itemId) & ")" & ODataQueryToString(query))
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
Public Function ItemsGet(Optional ByVal query As Variant) As stdJSON
  Dim outArr As stdJSON
  Dim nextUrl As String
  Dim pageJson As stdJSON
  Dim d As stdJSON
  Dim rows As stdJSON
  Dim i As Long

  Set outArr = stdJSON.Create(eJSONArray)
  nextUrl = ListItemsEndpoint(ODataQueryToString(query))

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
Public Function ItemAdd(ByVal data As stdJSON, Optional ByVal itemType As String = vbNullString) As Long
  Dim postData As stdJSON
  Dim Response As stdJSON
  Dim d As stdJSON
  Dim headers As Variant

  If data Is Nothing Then Err.Raise 5, "stdSharepointList::ItemAdd", "Data cannot be Nothing."

  Set postData = TransformSetterData(data)
  If LenB(itemType) = 0 Then itemType = GetListItemEntityType()
  SetMetadataType postData, itemType

  headers = BuildDefaultHeaders(True, SP_JSON_VERBOSE)
  stdHTTP.SetHeader headers, "X-RequestDigest", GetFormDigest()
  Set Response = RequestJson(ListItemsEndpoint(vbNullString), "POST", postData.ToString(), headers, SP_JSON_VERBOSE)
  Set d = JsonGetObject(Response, "d")
  If d Is Nothing Or Not d.Exists("Id") Then
    Err.Raise 5, "stdSharepointList::ItemAdd", "SharePoint did not return created item id."
  End If
  ItemAdd = CLng(d("Id"))
End Function

'Update (MERGE) a list item.
'@param ItemId - SharePoint item id.
'@param Data as stdJSON<Object> - Partial item payload.
'@returns as stdJSON - Response summary (`status`, `statusText`, optional `data`).
Public Function ItemSet(ByVal itemId As Variant, ByVal data As stdJSON) As stdJSON
  Dim PatchData As stdJSON
  Dim headers As Variant
  Dim req As stdHTTP
  Dim out As stdJSON

  If data Is Nothing Then Err.Raise 5, "stdSharepointList::ItemSet", "Data cannot be Nothing."

  Set PatchData = TransformSetterData(data)
  SetMetadataType PatchData, GetListItemEntityType()

  headers = BuildDefaultHeaders(True, SP_JSON_VERBOSE)
  stdHTTP.SetHeader headers, "X-RequestDigest", GetFormDigest()
  stdHTTP.SetHeader headers, "IF-MATCH", "*"
  stdHTTP.SetHeader headers, "X-HTTP-Method", "MERGE"

  Set req = SendRequest(ListItemsEndpoint("(" & CStr(itemId) & ")"), "POST", PatchData.ToString(), headers, SP_JSON_VERBOSE)
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

'Get the currently authenticated SharePoint user's email address.
'@returns - Current user email.
Public Function GetCurrentUserEmail() As String
  Dim json As stdJSON
  Dim d As stdJSON
  Dim loginName As String
  Dim claimSepPos As Long

  Set json = RequestJson(This.SiteURL & "/_api/web/currentuser?$select=Email,EMail,LoginName", "GET")
  Set d = JsonGetObject(json, "d")
  If d Is Nothing Then
    Err.Raise 5, "stdSharepointList::GetCurrentUserEmail", "SharePoint currentuser response did not contain 'd'."
  End If

  If d.Exists("Email") Then GetCurrentUserEmail = Trim$(CStr(d("Email")))
  If LenB(GetCurrentUserEmail) = 0 And d.Exists("EMail") Then GetCurrentUserEmail = Trim$(CStr(d("EMail")))
  If LenB(GetCurrentUserEmail) > 0 Then Exit Function

  If d.Exists("LoginName") Then
    loginName = Trim$(CStr(d("LoginName")))
    claimSepPos = InStrRev(loginName, "|")
    If claimSepPos > 0 Then loginName = Mid$(loginName, claimSepPos + 1)
    If InStr(1, loginName, "@", vbTextCompare) > 0 Then
      GetCurrentUserEmail = loginName
      Exit Function
    End If
  End If

  Err.Raise 5, "stdSharepointList::GetCurrentUserEmail", "SharePoint currentuser did not return an email. Use LoginName/UPN explicitly."
End Function

'Resolve/ensure a SharePoint user id from login/email text.
'@param UserText - User login, email, or principal text.
'@returns - SharePoint user id.
Public Function GetOrCreateUserId(ByVal UserText As String) As Long
  Dim key As String
  Dim body As String
  Dim json As stdJSON
  Dim d As stdJSON
  Dim headers As Variant

  key = LCase$(Trim$(UserText))
  If LenB(key) = 0 Then Err.Raise 5, "stdSharepointList::GetOrCreateUserId", "UserText cannot be blank."

  If This.UserIdCache.Exists(key) Then
    GetOrCreateUserId = CLng(This.UserIdCache(key))
    Exit Function
  End If

  body = "{""logonName"":" & JsonStringLiteral(UserText) & "}"
  headers = BuildDefaultHeaders(True, SP_JSON_VERBOSE)
  stdHTTP.SetHeader headers, "X-RequestDigest", GetFormDigest()
  Set json = RequestJson(This.SiteURL & "/_api/web/ensureuser", "POST", body, headers, SP_JSON_VERBOSE)
  Set d = JsonGetObject(json, "d")
  If d Is Nothing Or Not d.Exists("Id") Then
    Err.Raise 5, "stdSharepointList::GetOrCreateUserId", "SharePoint ensureuser did not return Id."
  End If

  This.UserIdCache.Add key, CLng(d("Id"))
  GetOrCreateUserId = CLng(d("Id"))
End Function

 'Resolve the uploader's distinct logins into site user IDs or per-login error messages.
Friend Sub BatchEnsureUsers(ByVal logins As Collection, ByVal people As Object, ByVal retryable As Collection)
  Const groupSize As Long = 50
  Dim first As Long, last As Long
  For first = 1 To logins.count Step groupSize
    last = first + groupSize - 1
    If last > logins.count Then last = logins.count
    EnsureUserGroup logins, first, last, people, retryable
  Next
End Sub

Private Sub EnsureUserGroup(ByVal logins As Collection, ByVal first As Long, ByVal last As Long, ByVal people As Object, ByVal retryable As Collection)
  On Error GoTo groupFailed
  Dim batch As String: batch = "batch_" & GenerateUuid()
  Dim lines As Collection: Set lines = New Collection
  Dim i As Long, change As String
  For i = first To last
    'One changeset per login keeps outer response order unambiguous, including aliases and failures.
    change = "changeset_" & Mid$(batch, 7) & "_" & CStr(i)
    lines.Add "--" & batch
    lines.Add "Content-Type: multipart/mixed; boundary=" & change
    lines.Add vbNullString
    lines.Add "--" & change
    lines.Add "Content-Type: application/http"
    lines.Add "Content-Transfer-Encoding: binary"
    lines.Add vbNullString
    lines.Add "POST " & This.SiteURL & "/_api/web/ensureuser HTTP/1.1"
    lines.Add "Content-Type: " & SP_JSON_VERBOSE
    lines.Add "Accept: " & SP_JSON_VERBOSE
    lines.Add vbNullString
    Dim body As stdJSON: Set body = stdJSON.Create(eJSONObject)
    body.Add "logonName", logins(i)
    lines.Add body.ToString()
    lines.Add vbNullString
    lines.Add "--" & change & "--"
  Next
  lines.Add "--" & batch & "--"
  lines.Add vbNullString
  Dim headers As Variant: headers = BuildDefaultHeaders(True, "multipart/mixed; boundary=" & batch)
  stdHTTP.SetHeader headers, "X-RequestDigest", GetFormDigest()
  Dim req As stdHTTP: Set req = SendRequest(This.SiteURL & "/_api/$batch", "POST", JoinCollection(lines, vbCrLf), headers, "multipart/mixed")
  Dim parts As Collection: Set parts = ParseODataBatchResponse(CStr(req.ResponseText))
  If parts.count <> last - first + 1 Then Err.Raise 5, "BatchEnsureUsers", "Unexpected ensureuser response count."
  For i = first To last
    Dim definitive As Boolean
    people(logins(i)) = EnsureUserResult(parts(i - first + 1), definitive)
    If Not definitive Then retryable.Add logins(i)
  Next
  Exit Sub
groupFailed:
  Dim detail As String: detail = "User lookup batch failed: " & Err.Description
  For i = first To last
    people(logins(i)) = detail
    retryable.Add logins(i)
  Next
End Sub

Private Function EnsureUserResult(ByVal part As Object, ByRef definitive As Boolean) As Variant
  definitive = False
  On Error GoTo invalidResult
  Dim status As Long: status = part("statusCode")
  Dim json As stdJSON: Set json = part("json")
  If status >= 200 And status < 300 Then
    Dim id As Long: id = VariantToLong(JsonGetObject(json, "d"), "Id")
    If id <= 0 Then GoTo invalidResult
    EnsureUserResult = id
    definitive = True
  Else
    Dim detail As String
    detail = JsonGetString(JsonGetObject(JsonGetObject(json, "error"), "message"), "value")
    EnsureUserResult = Trim$("HTTP " & status & " " & part("statusText") & " " & detail)
    definitive = (status = 400)
  End If
  Exit Function
invalidResult:
  EnsureUserResult = "Invalid ensureuser response (HTTP " & status & ")."
End Function

'Search SharePoint users by display/login/email text using People Picker.
'@param query - Search text. Case-insensitive contains matching is enforced post-filter.
'@param maxResults - Maximum number of candidates to return (clamped to safe bounds).
'@returns as stdJSON<Array> - Normalized user candidate objects.
Public Function SearchUsersByName(ByVal query As String, Optional ByVal maxResults As Long = 20) As stdJSON
  Dim normalizedQuery As String
  Dim requestedMax As Long
  Dim payload As stdJSON
  Dim queryParams As stdJSON
  Dim meta As stdJSON
  Dim headers As Variant
  Dim Response As stdJSON
  Dim rawResults As String
  Dim rows As stdJSON
  Dim outArr As stdJSON
  Dim one As stdJSON
  Dim i As Long
  Dim uniqueKey As String
  Dim seen As Object

  normalizedQuery = Trim$(query)
  If LenB(normalizedQuery) = 0 Then
    Err.Raise 5, "stdSharepointList::SearchUsersByName", "query cannot be blank."
  End If

  requestedMax = maxResults
  If requestedMax < 1 Then requestedMax = 1
  If requestedMax > PEOPLE_PICKER_MAX_RESULTS_CAP Then requestedMax = PEOPLE_PICKER_MAX_RESULTS_CAP

  Set payload = stdJSON.Create(eJSONObject)
  Set queryParams = stdJSON.Create(eJSONObject)
  Set meta = stdJSON.Create(eJSONObject)
  meta.Add "type", "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"
  queryParams.Add "__metadata", meta
  queryParams.Add "QueryString", normalizedQuery
  queryParams.Add "MaximumEntitySuggestions", requestedMax
  queryParams.Add "AllowEmailAddresses", True
  queryParams.Add "AllUrlZones", True
  queryParams.Add "PrincipalType", 1
  queryParams.Add "PrincipalSource", 15
  payload.Add "queryParams", queryParams

  headers = BuildDefaultHeaders(True, SP_JSON_VERBOSE)
  stdHTTP.SetHeader headers, "X-RequestDigest", GetFormDigest()
  Set Response = RequestJson( _
    This.SiteURL & "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser", _
    "POST", _
    payload.ToString(), _
    headers, _
    SP_JSON_VERBOSE _
  )

  Set outArr = stdJSON.Create(eJSONArray)
  rawResults = ExtractPeoplePickerResultJson(Response)
  If LenB(rawResults) = 0 Then
    Set SearchUsersByName = outArr
    Exit Function
  End If

  Set rows = TryParseJsonArray(rawResults)
  If rows Is Nothing Then
    Set SearchUsersByName = outArr
    Exit Function
  End If

  Set seen = NewDictionary(vbTextCompare)
  For i = 1 To rows.Length
    If IsObject(rows(i)) And TypeOf rows(i) Is stdJSON Then
      Set one = NormalizePeoplePickerCandidate(rows(i))
      If Not one Is Nothing Then
        If PeoplePickerCandidateMatches(one, normalizedQuery) Then
          uniqueKey = BuildPeoplePickerCandidateKey(one)
          If Not seen.Exists(uniqueKey) Then
            seen.Add uniqueKey, True
            outArr.Add one
            If outArr.Length >= requestedMax Then Exit For
          End If
        End If
      End If
    End If
  Next

  Set SearchUsersByName = outArr
End Function

'Batch update list items.
'@param BatchChanges as stdJSON<Array<{id: long, data: stdJSON}>> - Update operations.
'@param ChangeSetId - Optional changeset guid.
'@param BatchUuid - Optional batch guid.
'@param VerifyTimeouts - Read back requested fields by ID after a batch or item timeout.
'@returns as stdJSON<Array> - Per-item results. A verified timeout also carries verified=True.
Public Function BatchItemsSet( _
  ByVal BatchChanges As stdJSON, _
  Optional ByVal ChangeSetId As String = vbNullString, _
  Optional ByVal BatchUuid As String = vbNullString, _
  Optional ByVal VerifyTimeouts As Boolean = False _
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
  Dim Item As stdJSON
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
  For i = 1 To chunked.count
    Set chunk = chunked(i)
    If chunk.Length > 0 Then chunk.Add chunk(1) 'SP workaround parity with TypeScript implementation

    Set lineItems = New Collection
    lineItems.Add "--batch_" & BatchUuid
    lineItems.Add "Content-Type: multipart/mixed; boundary=changeset_" & ChangeSetId
    lineItems.Add vbNullString

    Dim wirePayloads As Collection: Set wirePayloads = New Collection
    For j = 1 To chunk.Length
      Set Item = chunk(j)
      Set op = TransformSetterData(JsonGetObject(Item, "data"))
      SetMetadataType op, itemType
      If VerifyTimeouts Then wirePayloads.Add op

      lineItems.Add "--changeset_" & ChangeSetId
      lineItems.Add "Content-Type: application/http"
      lineItems.Add "Content-Transfer-Encoding: binary"
      lineItems.Add vbNullString
      lineItems.Add "PATCH " & ListItemsEndpoint("(" & CStr(Item("id")) & ")") & " HTTP/1.1"
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
    Dim timedOut As Boolean
    Set req = PostChangeset(JoinCollection(lineItems, vbCrLf), headers, VerifyTimeouts, timedOut)
    If Not timedOut Then
      Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
      'Some servers omit the workaround response. Both shapes cover all original PATCHes.
      If parseRows.count <> chunk.Length And parseRows.count <> chunk.Length - 1 Then
        Err.Raise 5, "stdSharepointList::BatchItemsSet", "Unexpected PATCH response count. Check item state before retrying."
      End If
    End If

    Dim pending As Collection: Set pending = New Collection
    For j = 1 To chunk.Length - 1
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "update"
      out.Add "id", CLng(chunk(j)("id"))
      Dim status As Long: status = 0
      If Not timedOut Then
        Set one = parseRows(j)
        status = one("statusCode")
        If IsObject(one("json")) Then out.Add "data", one("json")
      End If
      out.Add "httpStatus", status
      out.Add "payload", chunk(j)
      If VerifyTimeouts And (timedOut Or status = 408 Or status = 504) Then
        out.Add "timeoutReason", "Update timed out. Requested values could not be confirmed."
        pending.Add Array(out, wirePayloads(j))
      End If
      results.Add out
    Next
    VerifyPendingUpdates pending
  Next

  Set BatchItemsSet = results
End Function

'Batch create list items.
'@param BatchItems as stdJSON<Array<stdJSON>> - Create payloads.
'@param BatchUuid - Optional batch guid.
'@param RecoveryProperty - Optional REST text property for a per-create GUID. On timeout a unique match returns its id and verified=True.
'@returns as stdJSON<Array> - Per-item result objects.
Public Function BatchItemsCreate(ByVal BatchItems As stdJSON, Optional ByVal BatchUuid As String = vbNullString, Optional ByVal RecoveryProperty As String = vbNullString) As stdJSON
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
  For i = 1 To chunked.count
    Set chunk = chunked(i)

    Set lineItems = New Collection
    lineItems.Add "--batch_" & BatchUuid
    lineItems.Add "Content-Type: multipart/mixed; boundary=changeset_" & BatchUuid
    lineItems.Add vbNullString

    Dim keys As Collection: Set keys = New Collection
    For j = 1 To chunk.Length
      Set op = TransformSetterData(chunk(j))
      If LenB(RecoveryProperty) > 0 Then
        keys.Add NewRecoveryKey()
        op.Add RecoveryProperty, keys(j)
      End If
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
    Dim timedOut As Boolean
    Set req = PostChangeset(JoinCollection(lineItems, vbCrLf), headers, keys.count > 0, timedOut)
    If Not timedOut Then
      Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
      If parseRows.count <> chunk.Length Then
        Err.Raise 5, "stdSharepointList::BatchItemsCreate", "Unexpected create response count. Check for existing items before retrying."
      End If
    End If
    Dim pending As Object: Set pending = NewDictionary(vbTextCompare)
    For j = 1 To chunk.Length
      Set out = stdJSON.Create(eJSONObject)
      out.Add "type", "create"
      Dim status As Long: status = 0
      If Not timedOut Then
        Set one = parseRows(j)
        status = one("statusCode")
        If IsObject(one("json")) Then
          out.Add "data", one("json")
          Set d = JsonGetObject(one("json"), "d")
          If Not d Is Nothing Then If d.Exists("Id") Then out.Add "id", CLng(d("Id"))
        End If
      End If
      If keys.count > 0 And (timedOut Or status = 408 Or status = 504) Then
        out.Add "timeoutReason", "Create timed out. Outcome unknown. Recovery key " & keys(j)
        pending.Add keys(j), out
      End If
      out.Add "httpStatus", status
      out.Add "payload", chunk(j)
      results.Add out
    Next
    Dim filters As Collection: Set filters = New Collection
    Dim key As Variant
    For Each key In pending.Keys
      filters.Add RecoveryProperty & " eq '" & key & "'"
      If filters.count = 25 Then
        RecoverCreates RecoveryProperty, filters, pending
        Set filters = New Collection
      End If
    Next
    If filters.count > 0 Then RecoverCreates RecoveryProperty, filters, pending
  Next

  Set BatchItemsCreate = results
End Function

 'Look up timed-out creates. Confirm only unique matches after reading every page.
Private Sub RecoverCreates(ByVal field As String, ByVal filters As Collection, ByVal pending As Object)
  On Error GoTo unconfirmed
  Dim url As String: url = ListItemsEndpoint("?$select=Id," & stdHTTP.URIEncode(field) & _
    "&$filter=" & stdHTTP.URIEncode(JoinCollection(filters, " or ")))
  Dim ids As Object: Set ids = NewDictionary(vbTextCompare)
  Do While LenB(url) > 0
    Dim page As stdJSON: Set page = RequestJson(url, "GET")
    Dim d As stdJSON: Set d = page("d")
    Dim rows As stdJSON: Set rows = d("results")
    Dim i As Long, key As Variant
    For i = 1 To rows.Length
      key = CStr(rows(i)(field))
      If ids.Exists(key) Then
        ids(key) = 0
      Else
        ids.Add key, CLng(rows(i)("Id"))
      End If
    Next
    url = JsonGetString(d, "__next")
  Loop
  For Each key In ids.Keys
    If pending.Exists(key) Then
      Dim result As stdJSON: Set result = pending(key)
      If ids(key) > 0 Then
        result.Add "id", ids(key)
        result.Add "verified", True
      Else
        result("timeoutReason") = "Multiple items match create recovery key " & key & ". Outcome unknown."
      End If
    End If
  Next
unconfirmed:
  'An incomplete lookup leaves the original timeout and key on each affected result.
End Sub

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
  For i = 1 To chunked.count
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
    Set req = SendRequest(This.SiteURL & "/_api/$batch", "POST", JoinCollection(lineItems, vbCrLf), headers, "multipart/mixed")

    Set parseRows = ParseODataBatchResponse(CStr(req.ResponseText))
    For j = 1 To parseRows.count
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
Public Function SimplifyPayload(ByVal payload As stdJSON) As stdJSON
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

  If payload Is Nothing Then Exit Function
  If Not payload.Exists("type") Then
    Set SimplifyPayload = JsonClone(payload)
    Exit Function
  End If

  pType = LCase$(CStr(payload("type")))
  If pType <> "update" Then
    Set SimplifyPayload = JsonClone(payload)
    Exit Function
  End If

  Set payloadObj = JsonGetObject(payload, "payload")
  Set dataObj = JsonGetObject(payloadObj, "data")
  Set prevObj = JsonGetObject(payload, "previousData")
  If dataObj Is Nothing Or prevObj Is Nothing Then
    Set SimplifyPayload = JsonClone(payload)
    Exit Function
  End If

  Set changed = stdJSON.Create(eJSONObject)
  Set info = dataObj.ChildrenInfo
  For i = 1 To info.count
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
  payloadObj.Add "id", payload("payload")("id")
  payloadObj.Add "data", changed
  out.Add "payload", payloadObj

  Set SimplifyPayload = out
End Function

'Process a set of commit payloads through batch methods.
'@param Payloads as stdJSON<Array> - Commit payloads with `type` + `payload`.
'@returns as stdJSON<{update: stdJSON[], create: stdJSON[], delete: stdJSON[]}> - Batch operation results.
Public Function ProcessCommitPayloads(ByVal payloads As stdJSON) As stdJSON
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

  If Not payloads Is Nothing Then
    For i = 1 To payloads.Length
      Set payload = payloads(i)
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
  ListEndpoint = This.SiteURL & "/_api/web/lists/GetByTitle('" & EscapeODataLiteral(This.ListTitle) & "')" & suffix
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

  options = stdHTTP.CreateOptions(body, headers, timeoutMs:=TIMEOUT_MAX_MS)

  Set req = stdHTTP.Create(url, method, HTTPSync, options, This.Authenticator)
  EnsureSuccess req, "stdSharepointList::SendRequest"
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
Private Function ODataQueryToString(Optional ByVal query As Variant) As String
  Dim params As Collection
  Dim q As stdJSON
  Dim selectCsv As String
  Dim expandCsv As String
  Dim filterText As String
  Dim useDefaultExpand As Boolean

  Set params = New Collection
  useDefaultExpand = True

  If IsMissingOrEmpty(query) Then
    BuildDefaultSelectExpand selectCsv, expandCsv
    If LenB(selectCsv) > 0 Then params.Add "$select=" & selectCsv
    If LenB(expandCsv) > 0 Then params.Add "$expand=" & expandCsv
  ElseIf VarType(query) = vbString Then
    ODataQueryToString = NormalizeRawQuery(CStr(query))
    Exit Function
  ElseIf IsObject(query) And TypeOf query Is stdJSON Then
    Set q = query
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
Private Function TransformGetterItem(ByVal Item As stdJSON) As stdJSON
  Dim out As stdJSON
  Dim idx As Long
  Dim def As TFieldDef
  Dim fieldName As String
  Dim FieldType As Long
  Dim rawValue As Variant

  Set out = JsonClone(Item)

  For idx = 1 To This.FieldsCount
    def = This.fields(idx)
    fieldName = def.InternalName
    FieldType = CLng(def.FieldType)
    If Not out.Exists(fieldName) Then GoTo ContinueField

    rawValue = out(fieldName)
    Select Case FieldType
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

  For i = 1 To info.count
    Set node = info(i)
    key = CStr(node("key"))
    lowerKey = LCase$(key)
    If IsObject(node("value")) Then
      Set val = node("value")
    Else
      Set val = Nothing
      val = node("value")
    End If

    'Preferred expanded-field alias syntax: Owner/Email
    slashPos = InStr(1, key, "/", vbTextCompare)
    If slashPos > 1 Then
      baseField = Left$(key, slashPos - 1)
      expandedProp = LCase$(Mid$(key, slashPos + 1))
      fieldIdx = FieldIndexByInternal(baseField)
      If fieldIdx > 0 Then
        def = This.fields(fieldIdx)
        If IsClearValue(val) Then
          If expandedProp = "email" Or expandedProp = "id" Then
            If CLng(def.FieldType) = SharePointFieldPerson Or CLng(def.FieldType) = SharePointFieldLookup Then
              JsonAddOrSet out, baseField & "Id", Null
              GoTo ContinueNode
            ElseIf CLng(def.FieldType) = SharePointFieldMultiPerson Then
              JsonAddOrSet out, baseField & "Id", BuildIdResultsObject(Null)
              GoTo ContinueNode
            End If
          End If
        End If
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
      def = This.fields(fieldIdx)
      If IsClearValue(val) Then
        Select Case CLng(def.FieldType)
          Case SharePointFieldPerson, SharePointFieldLookup
            JsonAddOrSet out, key & "Id", Null
          Case SharePointFieldMultiPerson
            JsonAddOrSet out, key & "Id", BuildIdResultsObject(Null)
          Case SharePointFieldMultiChoice
            JsonAddOrSet out, key, BuildStringResultsObject(Null)
          Case Else
            JsonAddOrSet out, key, Null
        End Select
      Else
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
            JsonAddOrSet out, key, FormatDateForSharePoint(val)
          Case Else
            JsonAddOrSet out, key, val
        End Select
      End If
    Else
      If IsClearValue(val) Then
        JsonAddOrSet out, key, Null
      Else
        JsonAddOrSet out, key, val
      End If
    End If
ContinueNode:
  Next

  Set TransformSetterData = out
End Function

'True when a setter value should clear the SharePoint field.
Private Function IsClearValue(ByVal value As Variant) As Boolean
  If IsNull(value) Then
    IsClearValue = True
  ElseIf IsEmpty(value) Then
    IsClearValue = True
  ElseIf VarType(value) = vbString Then
    IsClearValue = (LenB(Trim$(CStr(value))) = 0)
  End If
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

  Set json = RequestJson(This.SiteURL & "/_api/contextinfo", "POST", vbNullString, Empty, SP_JSON_VERBOSE)
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
    Err.Raise 5, "stdSharepointList::GetListItemEntityType", "ListItemEntityTypeFullName was not found."
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

 'Capture only timeouts from the write request, after payload preparation is complete.
Private Function PostChangeset(ByVal body As String, ByRef headers As Variant, ByVal captureTimeout As Boolean, ByRef timedOut As Boolean) As stdHTTP
  timedOut = False
  On Error GoTo failed
  Set PostChangeset = SendRequest(This.SiteURL & "/_api/$batch", "POST", body, headers, "multipart/mixed")
  Exit Function
failed:
  Dim number As Long: number = Err.Number
  Dim source As String: source = Err.source
  Dim detail As String: detail = Err.Description
  If (number = 10000 And source = "stdHTTP") Or number = &H80072EE2 Or _
      (source = "stdSharepointList::SendRequest" And (detail Like "HTTP 408 *" Or detail Like "HTTP 504 *")) Then
    timedOut = True
    If captureTimeout Then Exit Function
    source = "stdSharepointList::BatchWriteTimeout"
  End If
  Err.Raise number, source, detail
End Function

'Read a bounded group of IDs at a time. Each entry pairs the result with its original wire payload.
Private Sub VerifyPendingUpdates(ByVal pending As Collection)
  Const groupSize As Long = 25
  Dim first As Long, last As Long
  For first = 1 To pending.count Step groupSize
    last = first + groupSize - 1
    If last > pending.count Then last = pending.count
    VerifyUpdateGroup pending, first, last
  Next
End Sub

Private Sub VerifyUpdateGroup(ByVal pending As Collection, ByVal first As Long, ByVal last As Long)
  Dim byId As Object: Set byId = NewDictionary()
  Dim fields As Object: Set fields = NewDictionary()
  Dim filters As Collection: Set filters = New Collection
  Dim idx As Long, pair As Variant, child As Object, field As String, id As String
  Dim expected As stdJSON, result As stdJSON, actual As stdJSON
  On Error GoTo readFailed
  fields.Add "Id", True
  For idx = first To last
    pair = pending(idx)
    Set result = pair(0)
    id = CStr(result("id"))
    filters.Add "Id eq " & id
    byId.Add id, idx
    Set expected = pair(1)
    For Each child In expected.ChildrenInfo
      field = child("key")
      If field <> "__metadata" Then fields(field) = True
    Next
  Next

  Dim url As String
  url = ListItemsEndpoint("?$select=" & stdHTTP.URIEncode(Join(fields.Keys, ",")) & _
    "&$filter=" & stdHTTP.URIEncode(JoinCollection(filters, " or ")) & "&$top=" & CStr(last - first + 1))
  Do While LenB(url) > 0
    Dim page As stdJSON: Set page = RequestJson(url, "GET")
    Dim d As stdJSON: Set d = page("d")
    Dim rows As stdJSON: Set rows = d("results")
    Dim rowIndex As Long
    For rowIndex = 1 To rows.Length
      Set actual = rows(rowIndex)
      id = CStr(actual("Id"))
      If byId.Exists(id) Then
        pair = pending(byId(id))
        Set result = pair(0)
        Set expected = pair(1)
        If WrittenFieldsMatch(expected, actual) Then result.Add "verified", True
      End If
    Next
    url = JsonGetString(d, "__next")
  Loop
  Exit Sub
readFailed:
  Dim detail As String: detail = Err.Description
  For idx = first To last
    pair = pending(idx)
    Set result = pair(0)
    If Not result.Exists("verified") Then result("timeoutReason") = "Update timed out. Read-back failed: " & detail
  Next
End Sub

'Compare only submitted fields, using their wire names (including Person/Lookup Id).
Private Function WrittenFieldsMatch(ByVal expected As stdJSON, ByVal actual As stdJSON) As Boolean
  On Error GoTo unconfirmed
  Dim field As String, child As Object, compared As Boolean
  For Each child In expected.ChildrenInfo
    field = child("key")
    If field <> "__metadata" Then
      If Not actual.Exists(field) Then Exit Function
      If Not WrittenValueMatches(CStr(field), expected(field), actual(field)) Then Exit Function
      compared = True
    End If
  Next
  WrittenFieldsMatch = compared
unconfirmed:
End Function

Private Function WrittenValueMatches(ByVal field As String, ByVal wanted As Variant, ByVal found As Variant) As Boolean
  If IsNull(wanted) Or IsNull(found) Then
    WrittenValueMatches = IsNull(wanted) And IsNull(found)
  ElseIf IsObject(wanted) Or IsObject(found) Then
    If Not (IsObject(wanted) And IsObject(found)) Then Exit Function
    If Not (wanted.Exists("results") And found.Exists("results")) Then Exit Function
    Dim a As stdJSON: Set a = wanted("results")
    Dim b As stdJSON: Set b = found("results")
    If a.JsonType <> eJSONArray Or b.JsonType <> eJSONArray Or a.Length <> b.Length Then Exit Function
    Dim values As Object: Set values = NewDictionary()
    Dim i As Long, key As String
    For i = 1 To a.Length
      key = CollectionValueKey(a(i))
      values(key) = CLng(values(key)) + 1
    Next
    For i = 1 To b.Length
      key = CollectionValueKey(b(i))
      If Not values.Exists(key) Then Exit Function
      If values(key) = 1 Then values.Remove key Else values(key) = values(key) - 1
    Next
    WrittenValueMatches = values.count = 0
  ElseIf ResolveFieldTypeByName(field) = SharePointFieldDate Then
    Dim stamp As String: stamp = Replace(CStr(wanted), " ", "T")
    WrittenValueMatches = (CStr(found) = stamp & "Z" Or CStr(found) = stamp Or _
      CStr(found) = stamp & ".000Z")
  ElseIf VarType(wanted) = vbString Or VarType(found) = vbString Then
    If VarType(wanted) = VarType(found) Then WrittenValueMatches = StrComp(wanted, found, vbBinaryCompare) = 0
  ElseIf VarType(wanted) = vbBoolean Or VarType(found) = vbBoolean Then
    If VarType(wanted) = VarType(found) Then WrittenValueMatches = (wanted = found)
  ElseIf IsNumeric(wanted) And IsNumeric(found) Then
    WrittenValueMatches = (CDbl(wanted) = CDbl(found))
  End If
End Function

Private Function CollectionValueKey(ByVal value As Variant) As String
  If IsNumeric(value) And VarType(value) <> vbString Then
    CollectionValueKey = "number:" & CStr(value)
  Else
    CollectionValueKey = TypeName(value) & ":" & CStr(value)
  End If
End Function

'Parse multipart OData batch response into status/json result rows.
Private Function ParseODataBatchResponse(ByVal body As String) As Collection
  Set ParseODataBatchResponse = New Collection
  Dim boundary As String: boundary = ExtractBatchBoundary(body)
  If LenB(boundary) = 0 Then Err.Raise 5, "ParseODataBatchResponse", "No batch boundary found."
  ParseResponseParts Replace(body, vbCrLf, vbLf), boundary, ParseODataBatchResponse
End Function

'Only application/http MIME parts represent operations. Ignore the closing delimiter and epilogue.
Private Sub ParseResponseParts(ByVal body As String, ByVal boundary As String, ByVal rows As Collection)
  Dim parts() As String: parts = Split(vbLf & body, vbLf & boundary)
  Dim i As Long, part As String, pos As Long, headers As String, content As String
  Dim statusLine As String, tokens() As String, row As Object
  For i = 1 To UBound(parts)
    part = parts(i)
    If Left$(part, 2) = "--" Then Exit Sub
    If Left$(part, 1) <> vbLf Then GoTo malformed
    part = Mid$(part, 2)
    pos = InStr(part, vbLf & vbLf)
    If pos = 0 Then GoTo malformed
    headers = Left$(part, pos - 1)
    content = Mid$(part, pos + 2)
    pos = InStr(content & vbLf, vbLf)
    statusLine = Left$(content, pos - 1)
    If InStr(1, headers, "multipart/mixed", vbTextCompare) > 0 Then
      If Left$(statusLine, 2) <> "--" Or Len(statusLine) <= 2 Then GoTo malformed
      ParseResponseParts content, statusLine, rows
    Else
      If InStr(1, headers, "application/http", vbTextCompare) = 0 Then GoTo malformed
      tokens = Split(statusLine, " ", 3)
      If UBound(tokens) < 1 Then GoTo malformed
      If Left$(tokens(0), 5) <> "HTTP/" Or Not tokens(1) Like "###" Then GoTo malformed
      Set row = NewDictionary(vbTextCompare)
      row("statusCode") = CLng(tokens(1))
      row("statusText") = Mid$(statusLine, Len(tokens(0)) + Len(tokens(1)) + 2)
      Set row("json") = TryExtractJsonObject(content)
      rows.Add row
    End If
  Next
malformed:
  Err.Raise 5, "ParseODataBatchResponse", "Incomplete or malformed batch response. Check item state before retrying."
End Sub

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
  re.pattern = "--batchresponse_[\w\-]+"
  re.Global = False
  re.IgnoreCase = True
  If re.Test(body) Then
    Set m = re.Execute(body)(0)
    ExtractBatchBoundary = m.value
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
  Dim FieldType As Long
  FieldType = ResolveFieldTypeByName(key)

  If FieldType = SharePointFieldDate Then
    ValuesEqualForField = (NormalizeDateForCompare(leftValue) = NormalizeDateForCompare(rightValue))
  ElseIf IsObject(leftValue) Xor IsObject(rightValue) Then
    ValuesEqualForField = False
  ElseIf IsObject(leftValue) And IsObject(rightValue) Then
    If TypeOf leftValue Is stdJSON And TypeOf rightValue Is stdJSON Then
      ValuesEqualForField = (leftValue.ToString() = rightValue.ToString())
    Else
      ValuesEqualForField = (typeName(leftValue) = typeName(rightValue))
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
  ElseIf Not IsNull(value) Then
    If LenB(CStr(value)) > 0 Then results.Add CStr(value)
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
  ElseIf Not IsNull(value) Then
    If LenB(CStr(value)) > 0 Then results.Add ResolvePrincipalToUserId(value)
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
Private Function FieldIndexByInternal(ByVal InternalName As String) As Long
  Dim target As String
  Dim i As Long

  target = LCase$(Trim$(InternalName))
  If LenB(target) = 0 Then Exit Function

  For i = 1 To This.FieldsCount
    If LCase$(This.fields(i).InternalName) = target Then
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
    ResolveFieldTypeByName = CLng(This.fields(fieldIdx).FieldType)
    Exit Function
  End If

  slashPos = InStr(1, lower, "/", vbTextCompare)
  If slashPos > 1 Then
    candidate = Left$(lower, slashPos - 1)
    fieldIdx = FieldIndexByInternal(candidate)
    If fieldIdx > 0 Then
      ResolveFieldTypeByName = CLng(This.fields(fieldIdx).FieldType)
      Exit Function
    End If
  End If

  If Right$(lower, 2) = "id" Then
    candidate = Left$(lower, Len(lower) - 2)
    fieldIdx = FieldIndexByInternal(candidate)
    If fieldIdx > 0 Then
      ResolveFieldTypeByName = CLng(This.fields(fieldIdx).FieldType)
      Exit Function
    End If
  End If

  ResolveFieldTypeByName = SharePointFieldText
End Function

'Extract raw JSON array text from People Picker service response.
Private Function ExtractPeoplePickerResultJson(ByVal payload As stdJSON) As String
  Dim d As stdJSON

  If payload Is Nothing Then Exit Function
  Set d = JsonGetObject(payload, "d")
  If Not d Is Nothing Then
    If d.Exists("ClientPeoplePickerSearchUser") Then
      ExtractPeoplePickerResultJson = CStr(d("ClientPeoplePickerSearchUser"))
      Exit Function
    End If
  End If
  If payload.Exists("ClientPeoplePickerSearchUser") Then
    ExtractPeoplePickerResultJson = CStr(payload("ClientPeoplePickerSearchUser"))
  End If
End Function

'Normalize one People Picker entity into stable user candidate shape.
Private Function NormalizePeoplePickerCandidate(ByVal src As stdJSON) As stdJSON
  Dim out As stdJSON
  Dim entityData As stdJSON
  Dim displayText As String
  Dim email As String
  Dim loginName As String
  Dim userId As Long
  Dim principalType As Long

  If src Is Nothing Then Exit Function

  displayText = JsonGetString(src, "DisplayText")
  loginName = JsonGetString(src, "Key")
  If LenB(loginName) = 0 Then loginName = JsonGetString(src, "LoginName")
  email = JsonGetString(src, "Email")
  If LenB(email) = 0 Then email = JsonGetString(src, "EMail")

  Set entityData = JsonGetObject(src, "EntityData")
  If Not entityData Is Nothing Then
    If LenB(email) = 0 Then email = JsonGetString(entityData, "Email")
    If LenB(email) = 0 Then email = JsonGetString(entityData, "EMail")
    If LenB(loginName) = 0 Then loginName = JsonGetString(entityData, "PrincipalName")
    If LenB(loginName) = 0 Then loginName = JsonGetString(entityData, "AccountName")
    If LenB(displayText) = 0 Then displayText = JsonGetString(entityData, "Title")

    principalType = VariantToLong(entityData, "PrincipalType")
    If principalType > 0 And principalType <> 1 Then Exit Function

    userId = VariantToLong(entityData, "SPUserID")
    If userId = 0 Then userId = VariantToLong(entityData, "PrincipalId")
  End If

  If LenB(displayText) = 0 Then displayText = loginName
  If LenB(displayText) = 0 Then displayText = email
  If LenB(displayText) = 0 Then Exit Function

  Set out = stdJSON.Create(eJSONObject)
  out.Add "DisplayText", displayText
  out.Add "Email", email
  out.Add "LoginName", loginName
  If userId > 0 Then out.Add "UserId", userId

  Set NormalizePeoplePickerCandidate = out
End Function

'Check whether normalized candidate contains query in searchable fields.
Private Function PeoplePickerCandidateMatches(ByVal candidate As stdJSON, ByVal query As String) As Boolean
  If candidate Is Nothing Then Exit Function
  If InStr(1, JsonGetString(candidate, "DisplayText"), query, vbTextCompare) > 0 Then
    PeoplePickerCandidateMatches = True
    Exit Function
  End If
  If InStr(1, JsonGetString(candidate, "LoginName"), query, vbTextCompare) > 0 Then
    PeoplePickerCandidateMatches = True
    Exit Function
  End If
  If InStr(1, JsonGetString(candidate, "Email"), query, vbTextCompare) > 0 Then
    PeoplePickerCandidateMatches = True
  End If
End Function

'Build a stable key used to deduplicate candidate rows.
Private Function BuildPeoplePickerCandidateKey(ByVal candidate As stdJSON) As String
  Dim key As String
  key = LCase$(Trim$(JsonGetString(candidate, "LoginName")))
  If LenB(key) = 0 Then key = LCase$(Trim$(JsonGetString(candidate, "Email")))
  If LenB(key) = 0 Then key = LCase$(Trim$(JsonGetString(candidate, "DisplayText")))
  If LenB(key) = 0 Then key = "#unknown"
  BuildPeoplePickerCandidateKey = key
End Function

 'Build default schema-driven `$select` and `$expand` fragments.
Private Sub BuildDefaultSelectExpand(ByRef selectCsv As String, ByRef expandCsv As String)
  Dim selectParts As Collection
  Dim expandParts As Collection
  Dim fieldDef As TFieldDef
  Dim InternalName As String
  Dim fType As Long
  Dim props As Collection
  Dim i As Long
  Dim idx As Long

  Set selectParts = New Collection
  Set expandParts = New Collection
  AddUnique selectParts, stdHTTP.URIEncode("Id")

  For idx = 1 To This.FieldsCount
    fieldDef = This.fields(idx)
    InternalName = fieldDef.InternalName
    fType = CLng(fieldDef.FieldType)

    Select Case fType
      Case SharePointFieldPerson, SharePointFieldMultiPerson
        AddUnique expandParts, stdHTTP.URIEncode(InternalName)
        Set props = fieldDef.SelectProps
        If props.count = 0 Then
          AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/Id")
          AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/Title")
          AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/EMail")
        Else
          For i = 1 To props.count
            AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/" & CStr(props(i)))
          Next
        End If
      Case SharePointFieldLookup, SharePointFieldFile
        AddUnique expandParts, stdHTTP.URIEncode(InternalName)
        Set props = fieldDef.SelectProps
        If props.count = 0 Then
          AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/Id")
          AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/Title")
        Else
          For i = 1 To props.count
            AddUnique selectParts, stdHTTP.URIEncode(InternalName & "/" & CStr(props(i)))
          Next
        End If
      Case Else
        AddUnique selectParts, stdHTTP.URIEncode(InternalName)
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
  If params.count = 0 Then Exit Function
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
  For i = 1 To items.count
    If i > 1 Then sb = sb & delim
    sb = sb & CStr(items(i))
  Next
  JoinCollection = sb
End Function

 'Add a string value to collection only if not already present.
Private Sub AddUnique(ByVal list As Collection, ByVal value As String)
  Dim i As Long
  For i = 1 To list.count
    If CStr(list(i)) = value Then Exit Sub
  Next
  list.Add value
End Sub

 'Parse list page URL into site URL, list title, and relative path.
Private Sub ParseListUrl(ByVal raw As String, ByRef SiteURL As String, ByRef ListTitle As String, ByRef listRelativePath As String)
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
  If listsPos <= 0 Then Err.Raise 5, "stdSharepointList::ParseListUrl", "ListUrl must contain `/Lists/<ListName>`."

  SiteURL = Left$(cleaned, listsPos - 1)
  afterLists = Mid$(cleaned, listsPos + Len("/Lists/"))
  slashPos = InStr(1, afterLists, "/", vbTextCompare)
  If slashPos > 0 Then
    listPathSegment = Left$(afterLists, slashPos - 1)
  Else
    listPathSegment = afterLists
  End If

  ListTitle = stdHTTP.URIDecode(listPathSegment)
  SiteURL = TrimTrailingSlash(SiteURL)
  listRelativePath = BuildServerRelativeListPath(SiteURL, listPathSegment)
End Sub

Private Sub EnsureListDisplayTitle()
  Dim resolvedTitle As String

  If This.ListTitleResolved Then Exit Sub

  resolvedTitle = ResolveListDisplayTitle()
  If LenB(Trim$(resolvedTitle)) = 0 Then
    Err.Raise 5, "stdSharepointList::EnsureListDisplayTitle", "Unable to resolve list display title from list URL."
  End If

  This.ListTitle = resolvedTitle
  This.ListTitleResolved = True
End Sub

Private Function ResolveListDisplayTitle() As String
  Dim api As String
  Dim json As stdJSON
  Dim d As stdJSON

  api = This.SiteURL & "/_api/web/GetList('" & EscapeODataLiteral(This.listRelativePath) & "')?$select=Title"
  Set json = RequestJson(api, "GET")
  Set d = JsonGetObject(json, "d")

  If d Is Nothing Or Not d.Exists("Title") Then
    Err.Raise 5, "stdSharepointList::ResolveListDisplayTitle", "SharePoint did not return a list Title."
  End If

  ResolveListDisplayTitle = CStr(d("Title"))
End Function

Private Function BuildServerRelativeListPath(ByVal SiteURL As String, ByVal listPathSegment As String) As String
  Dim siteRelativePath As String
  Dim decodedSegment As String

  siteRelativePath = SiteServerRelativePath(SiteURL)
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

'Get a child value as a trimmed string, returning empty when missing.
Private Function JsonGetString(ByVal parent As stdJSON, ByVal key As String) As String
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then Exit Function
  If IsNull(parent(key)) Then Exit Function
  JsonGetString = Trim$(CStr(parent(key)))
End Function

'Get a child value as Boolean, returning False when missing/invalid.
Private Function JsonGetBoolean(ByVal parent As stdJSON, ByVal key As String) As Boolean
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then Exit Function
  If IsNull(parent(key)) Then Exit Function
  On Error Resume Next
  JsonGetBoolean = CBool(parent(key))
  On Error GoTo 0
End Function

'Convert a named JSON property to Long safely, returning 0 when absent/invalid.
Private Function VariantToLong(ByVal parent As stdJSON, ByVal key As String) As Long
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then Exit Function
  If Not IsNumeric(parent(key)) Then Exit Function
  VariantToLong = CLng(parent(key))
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
Private Function JsonClone(ByVal Source As stdJSON) As stdJSON
  If Source Is Nothing Then
    Set JsonClone = stdJSON.Create(eJSONObject)
  Else
    Set JsonClone = stdJSON.CreateFromVariant(Source.ToVBObject())
  End If
End Function

'Attempt to parse JSON text into a JSON array.
Private Function TryParseJsonArray(ByVal jsonText As String) As stdJSON
  jsonText = Trim$(jsonText)
  If LenB(jsonText) = 0 Then Exit Function

  On Error Resume Next
  Set TryParseJsonArray = stdJSON.CreateFromString(jsonText)
  On Error GoTo 0

  If TryParseJsonArray Is Nothing Then Exit Function
  If TryParseJsonArray.JsonType <> eJSONArray Then Set TryParseJsonArray = Nothing
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
    ElseIf typeName(SelectProps) = "Collection" Then
      For i = 1 To SelectProps.count
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
  NewDictionary.compareMode = compareMode
  On Error GoTo 0
End Function

 'Generate an OS-backed identifier that stays unique across concurrent uploaders.
Private Function NewRecoveryKey() As String
  Dim bytes(0 To 15) As Byte
  #If Mac Then
    uuid_generate_random bytes(0)
  #Else
    If CoCreateGuid(bytes(0)) <> 0 Then Err.Raise 5, "NewRecoveryKey", "Could not generate a create recovery GUID."
  #End If
  Dim i As Long, idx As Long, value As String
  For i = 0 To 15
    idx = i
    #If Not Mac Then
      Select Case i
        Case 0 To 3: idx = 3 - i
        Case 4 To 5: idx = 9 - i
        Case 6 To 7: idx = 13 - i
      End Select
    #End If
    Select Case i
      Case 4, 6, 8, 10: value = value & "-"
    End Select
    value = value & Right$("0" & Hex$(bytes(idx)), 2)
  Next
  NewRecoveryKey = LCase$(value)
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
