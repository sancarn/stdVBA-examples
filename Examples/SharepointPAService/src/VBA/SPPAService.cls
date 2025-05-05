VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SPPAService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum SharepointFileCheckInType
  CheckInMinor     'Check in as a draft (minor) version.
  CheckInMajor     'Check in as a published (major) version.
  CheckInOverwrite 'Overwrite the current version.
End Enum

Private Type TThis
  serviceWorker As String
  SiteURL As String
  listSelectorType As String
  listSelectorValue As String
  listAPI As String
  docLibAPI As String
End Type
Private This As TThis

'Create a power automate service worker to automate the creation, maintainence and deletion of sharepoint items
'@constructor
'@param serviceWorker - The url to the power automate service worker.
'@param siteURL - The sharepoint site url.
'@param listSelectorValue - title, id or name of the list to run the queries against.
'@param listSelectorType - "title", "id" or "name" depending on how you are wanting to select the list in question.
'@returns - Sharepoint PowerAutomate service worker.
'@example `set service = SPPAService.Create("http://service.worker.com", "http://tenant.sharepoint.com/sites/My/Site", "My List")`
Public Function Create(ByVal serviceWorker As String, ByVal SiteURL As String, ByVal listSelectorValue As String, Optional ByVal listSelectorType As String = "title") As SPPAService
  Set Create = New SPPAService
  Call Create.protInit(serviceWorker, SiteURL, listSelectorValue, listSelectorType)
End Function

'Initialisation sub
'@protected
'@param siteURL - The sharepoint site url
Public Sub protInit(serviceWorker As String, SiteURL As String, listSelectorValue As String, listSelectorType As String)
  This.serviceWorker = serviceWorker
  This.SiteURL = SiteURL
  This.listSelectorType = listSelectorType
  This.listSelectorValue = listSelectorValue
  
  Const APITemplate As String = "_api/web/lists${selector}"
  Select Case lcase(listSelectorType)
    Case "title"
      This.listAPI = InterpolateString(APITemplate, "selector", "/GetByTitle('${title}')", "title", listSelectorValue)
    Case "id"
      This.listAPI = InterpolateString(APITemplate, "selector", "guid('${guid}')", "guid", listSelectorValue)
    Case "name"
      This.listAPI = InterpolateString(APITemplate, "selector", "/" & listSelectorValue)
  End Select
End Sub


'Create a power automate service worker from a configuration file. These configuration files can be created with the `ToConfigFile` method.
'@constructor
'@param workerFile - Path to a worker config file. The file should be in the JSON format with a `url` property.
'@param listOrLibraryFile - Path to a list/library config file. The file should be in the JSON format with a `siteURL`, `selectorType` and `selectorValue` properties.
'@returns - Sharepoint PowerAutomate service worker.
'@example `set service = SPPAService.CreateFromConfigs("C:\Tmp\MyWorker.json", "C:\Tmp\MyList.json")`
Public Function CreateFromConfigs(ByVal workerFile As String, ByVal listOrLibraryFile As String) As SPPAService
  Dim workerConfig As stdJSON: Set workerConfig = stdJSON.CreateFromFile(workerFile)
  Dim llConfig As stdJSON: Set llConfig = stdJSON.CreateFromFile(listOrLibraryFile)
  Set CreateFromConfigs = Create(workerConfig.item("url"), llConfig.item("siteURL"), llConfig.item("selectorValue"), llConfig.item("selectorType"))
End Function


'Create a power automate service worker from a configuration file. These configuration files can be created with the `ToConfigFile` method.
'@constructor
'@param configFile - Path to a config file. The file should be in the JSON format with a `serviceWorker` and a `siteURL` property.
'@returns - Sharepoint PowerAutomate service worker.
'@example ```vba
'SPPAService.CreateFromCombinedConfig("C:\Tmp\MyConnectionFile.json")
'```
Public Function CreateFromCombinedConfig(ByVal configFile As String) As SPPAService
  Dim config As stdJSON: Set config = stdJSON.CreateFromFile(configFile)
  Set CreateFromCombinedConfig = Create(config.item("serviceWorker"), config.item("list").item("siteURL"), config.item("list").item("selectorValue"), config.item("list").item("selectorType"))
End Function


'Creates a worker config file from the existing instance
'@param configFile - Path to the desired config file location
Public Sub ToWorkerConfig(ByVal configFile As String)
  With stdJSON.Create()
    .Add "url", This.serviceWorker
    Call .ToFile(configFile, 2)
  End With
End Sub

'Creates a list/library config file from the existing instance
'@param configFile - Path to the desired config file location
Public Sub ToListLibConfig(ByVal configFile As String)
  With stdJSON.Create()
    .Add "siteURL", This.SiteURL
    .Add "selectorValue", This.listSelectorValue
    .Add "selectorType", This.listSelectorType
    Call .ToFile(configFile, 2)
  End With
End Sub

'Creates a config file from the existing instance
'@param configFile - Path to the config file to be produced
Public Sub ToCombinedConfig(ByVal configFile As String)
  With stdJSON.Create()
    .Add "serviceWorker", This.serviceWorker
    With .AddObject("list")
      .Add "siteURL", This.SiteURL
      .Add "selectorValue", This.listSelectorValue
      .Add "selectorType", This.listSelectorType
    End With
    Call .ToFile(configFile, 2)
  End With
End Sub

Public Property Get SiteURL() As String
  SiteURL = This.SiteURL
End Property

Public Property Get ListSelector() As String
  ListSelector = This.listSelectorValue
End Property

'Execute an arbitrary query using the Sharepoint Service
'@param apiPath -
'@param method -
'@param headers -
'@param body -
'@returns -
'@example ```
'Dim sp as SPPAService: set sp = SPPAService.CreateFromConfigFile("C:\configs\SPPA_MySite.json")
'Dim http as stdHTTP: set http = sp.RawRequest( _
  "_api/_api/web/lists/GetByTitle('My data')", _
  "GET", _
  stdJSON.CreateFromParams(eJSONObject,"Accept", "application/json;odata=verbose") _
)
'Debug.Print http.ResponseText
'```
'@remarks All responses have a `ResponseMapper` to parse the response text as JSON. This may need to be overriden however.
Public Function RawRequest(ByVal apiPath As String, ByVal method As String, ByVal headers As stdJSON, Optional ByVal body As String = "") As stdHTTP
  With stdJSON.Create()
    .Add "site", This.SiteURL
    .Add "api", apiPath
    .Add "method", method
    .Add "headers", headers
    .Add "body", body

    Dim reqBody As String: reqBody = .ToString()
    
    'For Debugging
    'Dim asReq As String
    'asReq = method & " " & This.SiteURL & "/" & apiPath & " HTTP/1.1" & vbCrLf
    'asReq = asReq & headers.ToString(True) & vbCrLf
    'asReq = asReq & vbCrLf
    'asReq = asReq & body & vbCrLf
    'stdClipboard.text = asReq
    

    Set RawRequest = stdHTTP.Create( _
      This.serviceWorker, _
      "POST", _
      ThreadingStyle:=HTTPAsync, _
      options:=stdHTTP.CreateOptions(reqBody, TimeoutMS:=300000) _
    )

    'Always assume JSON response, but allow override
    Set RawRequest.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText)").Bind(stdJSON)
  End With
End Function

'Get a single item from a list by it's id
'@param itemID - The ID of the item to retrieve
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItem(ByVal itemID As Long) As stdHTTP
  Const APITemplate As String = "${list}/items(${itemID})"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItem = RawRequest(API, "GET", headers)
End Function

'Create a single item in a list
'@param data - The data the new item should contain
'@param itemType - Optional. If not supplied it will be queried at runtime on every call. Query once to increase performance, or add as a known value. Can be obtained using `getListItemType` function
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemCreate(ByVal data As stdJSON, Optional ByVal itemType As String = "") As stdHTTP
  'Prepare item type
  If itemType = "" Then itemType = getListItemType().Await().ResponseMapped()
  With data.AddObject("__metadata")
    .Add "type", itemType
  End With

  Const APITemplate As String = "${list}/items"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI)
  Dim headers As stdJSON
  With stdJSON.Create()
    .Add "Accept", "application/json;odata=verbose"
    .Add "Content-Type", "application/json;odata=verbose"
    Set headers = .ToSelf()
  End With

  Set ListItemCreate = RawRequest(API, "POST", headers, data.ToString())
End Function

'Update a single item in a list
'@param itemID - The ID of the item to update
'@param data - The data the new item should contain
'@param itemType - Optional. If not supplied it will be queried at runtime on every call. Query once to increase performance, or add as a known value. Can be obtained using `getListItemType` function
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemUpdate(ByVal itemID As Long, ByVal data As stdJSON, Optional ByVal itemType As String = "") As stdHTTP
  'Prepare item type
  If itemType = "" Then itemType = getListItemType().Await().ResponseMapped()
  With data.AddObject("__metadata")
    .Add "type", itemType
  End With

  Const APITemplate As String = "${list}/items(${itemID})"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID)
  Dim headers As stdJSON
  With stdJSON.Create()
    .Add "Accept", "application/json;odata=verbose"
    .Add "Content-Type", "application/json;odata=verbose"
    .Add "If-Match", "*"
    .Add "X-HTTP-Method", "MERGE"
    Set headers = .ToSelf()
  End With

  Set ListItemUpdate = RawRequest(API, "POST", headers, data.ToString())
End Function

'Delete a single item in a list
'@param itemID - The ID of the item to delete
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemDelete(ByVal itemID As Long) As stdHTTP
  Const APITemplate As String = "${list}/items(${itemID})"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID)
  Dim headers As stdJSON
  With stdJSON.Create()
    .Add "Accept", "application/json;odata=verbose"
    .Add "Content-Type", "application/json;odata=verbose"
    .Add "If-Match", "*"
    .Add "X-HTTP-Method", "DELETE"
    Set headers = .ToSelf()
  End With

  Set ListItemDelete = RawRequest(API, "POST", headers)
End Function

'Get all attachments assosciated with a list item
'@param itemID - The ID of the item to get the attachments for
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
'@docs https://chatgpt.com/share/67fceb72-7f5c-8013-8208-7143a5539e30
Public Function ListItemAttachmentsGet(ByVal itemID As Long) As stdHTTP
  Const APITemplate As String = "${list}/items(${item})/AttachmentFiles"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItemAttachmentsGet = RawRequest(API, "GET", headers)
End Function

'Get a specific attachment by name from a list item
'@param itemID - The ID of the item to get the attachment for
'@param fileName - The name of the file to retrieve
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
'@docs https://chatgpt.com/share/67fceb72-7f5c-8013-8208-7143a5539e30
Public Function ListItemAttachmentGet(ByVal itemID As Long, ByVal fileName As String) As stdHTTP
  Const APITemplate As String = "${list}/items(${item})/AttachmentFiles('${fileName}')/$value"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID, "fileName", fileName)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItemAttachmentGet = RawRequest(API, "GET", headers)
End Function

'Create an attachment in a sharepoint list
'@param itemID - The ID of the item to create an attachment within
'@param fileName - The name of the file to create
'@param contents - The contents of the file to create
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
'@docs https://chatgpt.com/share/67fceb72-7f5c-8013-8208-7143a5539e30
Public Function ListItemAttachmentCreate(ByVal itemID As Long, ByVal fileName As String, contents() As Byte) As stdHTTP
  Const APITemplate As String = "${list}/items(${item})/AttachmentFiles/add(FileName='${fileName}')"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID, "fileName", fileName)
  'TODO: Do we need a X-RequestDigest?
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItemAttachmentCreate = RawRequest(API, "POST", headers, contents)
End Function

'Update an attachment in a sharepoint list
'@param itemID - The ID of the item to update the attachment of
'@param fileName - The name of the file to update
'@param contents - The new contents of the file to update
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
'@docs https://chatgpt.com/share/67fceb72-7f5c-8013-8208-7143a5539e30
Public Function ListItemAttachmentUpdate(ByVal itemID As Long, ByVal fileName As String, contents() As Byte) As stdHTTP
  Const APITemplate As String = "${list}/items(${item})/AttachmentFiles('${fileName}')/$value"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID, "fileName", fileName)
  'TODO: Do we need a X-RequestDigest?
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "X-HTTP-Method", "PUT")

  Set ListItemAttachmentUpdate = RawRequest(API, "POST", headers, contents)
End Function

'Delete an attachment in a sharepoint list
'@param itemID - The ID of the item to delete the attachment
'@param fileName - The name of the file to delete
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest
'@docs https://chatgpt.com/share/67fceb72-7f5c-8013-8208-7143a5539e30
Public Function ListItemAttachmentDelete(ByVal itemID As Long, ByVal fileName As String) As stdHTTP
  Const APITemplate As String = "${list}/items(${item})/AttachmentFiles('${fileName}')"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "itemID", itemID, "fileName", fileName)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItemAttachmentUpdate = RawRequest(API, "DELETE", headers)
End Function


'Get all list items from a sharepoint list
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItems(Optional ByVal query As String) As stdHTTP
  Const APITemplate As String = "${list}/items?${query}"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "query", query)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItems = RawRequest(API, "GET", headers)
End Function

'Get list items filtering to an ODATA Query
'@example `myList.ListItemsFromODataQuery("myCol eq 'test'")`
'@remark ODATA Queries do not work effectively on lists with more than 500 records. It's recommended to use `ListItemsFromInplaceQuery` instead where possible.
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
Public Function ListItemsFromODataQuery(ByVal query As String) As stdJSON
  Const APITemplate As String = "${list}/items?$filter=${oDataQuery}"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "oDataQuery", query)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")
  Dim data As stdHTTP: Set data = RawRequest(API, "GET", headers).Await().ResponseMapped().item("d")
  Dim result As stdJSON: Set result = stdJSON.Create(eJSONArray)
  Call result.Add(data.item("results"))

  While Not isNull(data.item("__next"))
    Set data = RawRequest(data.item("__next"), "GET", headers).Await().ResponseMapped().item("d")
    Call result.Add(data.item("results"))
  Wend

  Set ListItemsFromODataQuery = result
End Function

'Get all items using an inplace query
Public Function ListItemsFromInplaceQuery(ByVal query As String) As stdHTTP
  Const APITemplate As String = "${list}/RenderListDataAsStream?InplaceSearchQuery=${inplaceQuery}"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI, "inplaceQuery", query)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set ListItemsFromInplaceQuery = RawRequest(API, "POST", headers)
  Set ListItemsFromInplaceQuery.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText).Row").Bind(stdJSON)
End Function

'Uses $batch query to add many items to a sharepoint list in as few requests as possible
'@param additions as stdArray<stdJSON>- Array of JSON objects in the form `{field1: ..., ...}`
'@param changeSetID - Optional. If not supplied it will be generated at runtime.
'@param batchUUID - Optional. If not supplied it will be generated at runtime.
'@returns - Collection of the responses from the batch requests. Call `ResponseMapped()` to get the status and data of all requests.
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsCreateBatch(ByVal additions As stdArray, Optional ByVal changeSetID As String, Optional ByVal batchUUID As String) As HTTPCollection
  Dim itemType As String: itemType = getListItemType().Await.ResponseMapped()
  If changeSetID = "" Then changeSetID = getGUID()
  If batchUUID = "" Then batchUUID = getGUID()
  Dim responses As HTTPCollection: Set responses = HTTPCollection.Create()
  
  'A max of 1000 operations are allowed in a changeset; To stay well under this value, we use batches of 750.
  For Each AddSet In additions.SplitToChunks(750)
    '2025-04-15 Doesn't appear to be an issue with additions
    ''Workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
    'Call AddSet.Push(AddSet.item(1))

    'Create addSet
    Dim batch As stdArray: Set batch = stdArray.Create()
    Call batch.Push("--batch_" & batchUUID)
    Call batch.Push("Content-Type: multipart/mixed; boundary=changeset_" & changeSetID)
    Call batch.Push("")

    'Metadata will be used in each step
    Dim metadata As stdJSON: Set metadata = stdJSON.CreateFromParams(eJSONObject, "type", itemType)

    'Each item of the form: {"Field1": Value1, ...}
    Dim item As stdJSON
    For Each item In AddSet
      'Bind metadata
      If Not item.Exists("__metadata") Then Call item.Add("__metadata", metadata)
      
      'Add change
      Call batch.Push("--changeset_" & changeSetID)
      Call batch.Push("Content-Type: application/http")
      Call batch.Push("Content-Transfer-Encoding: binary")
      Call batch.Push("")
      Call batch.Push(InterpolateString("POST ${site}/${list}/items HTTP/1.1", "site", This.SiteURL, "list", This.listAPI))
      Call batch.Push("Content-Type: application/json;odata=verbose")
      Call batch.Push("Accept: application/json;odata=verbose")
      Call batch.Push("")
      Call batch.Push(item.ToString())
      Call batch.Push("")
    Next

    'End changeset to create Data
    Call batch.Push("--changeset_" & changeSetID & "--")
    Call batch.Push("--batch_" & batchUUID & "--")
    
    Dim body As String: body = batch.Join(vbCrLf)

    'Make request via RawRequest
    Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Content-Type", "multipart/mixed; boundary=batch_" & batchUUID, "Accept", "application/json")
    Dim request As stdHTTP: Set request = RawRequest("_api/$batch", "POST", headers, body)
    Call responses.add(request)
  Next

  Set ListItemsCreateBatch = responses
  set ListItemsCreateBatch.responseMapper = stdCallback.CreateFromObjectMethod(Me, "protProcessBatchResponse").bind("addition", additions)
End Function

'Process the batch response from a batch request
'@protected
'@param typeOfRequest - "create", "update" or "delete" dependent on which request was made
'@param initialData as stdArray<stdJSON> - The data that was sent in the request. This is used to map the response to the original request.
'@param col as Collection<stdHTTP> - The collection of batch HTTP requests/responses to process
'@returns stdArray<stdJSON> - A collection of responses from the batch request
Public Function protProcessBatchResponse(ByVal typeOfRequest as string, ByVal initialData as stdArray, ByVal col as Collection) as stdArray
  Dim regex as stdRegex: set regex = stdRegex.Create("HTTP\/1.1 (?<status>\d+) (?<statusText>.*)\s+(?<headers>(?:.+: .+\n)*)\n(?<body>.*)")
  Dim responses as stdArray: Set responses = stdArray.CreateWithOptions(initialData.length, 10)
  Dim index as long: index = 0
  Dim request As stdHTTP
  For each request in col
    'Await request to ensure it is complete
    Call request.Await()

    'Parse the response
    Dim batchResponse as string: batchResponse = request.ResponseText
    Dim cMatches As Collection: Set matches = regex.MatchAll(batchResponse)
    Dim aMatches as stdArray: set aMatches = stdArray.CreateFromCollection(matches)
    
    'Remove last item as per workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
    select case typeOfRequest
      case "update", "delete"
        set aMatches = aMatches.Slice(0,-1) 
    end select
    
    'Loop matches and zip with original data
    Dim match As object
    For each match in aMatches
      index = index + 1
      With stdJSON.Create()
        .add "status", match.item("status")
        .add "statusText", match.item("statusText")
        .add "headers", stdJSON.CreateFromString(match.item("headers"))
        .add "body", stdJSON.CreateFromString(match.item("body"))
        .add "input", initialData.item(index)
        Call responses.Push(.ToSelf())
      end with
    Next
  Next
  set protProcessBatchResponse = responses
End Function

'Uses $batch query to update many items in a sharepoint list in as few requests as possible
'@param changes - Changes in the form `{id: ..., data: {field1: ..., ...}}`
'@param changeSetID - an id for each sub-request
'@param batchUUID - an id for the whole request
'@returns - HTTPCollection of responses. Use `http.ResponseMapped()` to get returned data
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsUpdateBatch(ByVal changes As stdArray, Optional ByVal changeSetID As String, Optional ByVal batchUUID As String) As HTTPCollection
  Dim itemType As String: itemType = getListItemType().Await().ResponseMapped()
  If changeSetID = "" Then changeSetID = getGUID()
  If batchUUID = "" Then batchUUID = getGUID()
  Dim responses As HTTPCollection: Set responses = HTTPCollection.Create()

  'A max of 1000 operations are allowed in a changeset; To stay well under this value, we use batches of 750.
  For Each changeSet In changes.SplitToChunks(750)
    'Workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
    Call changeSet.Push(changeSet.item(1))

    'Create changeset
    Dim batch As stdArray: Set batch = stdArray.Create()
    Call batch.Push("--batch_" & batchUUID)
    Call batch.Push("Content-Type: multipart/mixed; boundary=changeset_" & changeSetID)
    Call batch.Push("")

    'Metadata will be used in each step
    Dim metadata As stdJSON: Set metadata = stdJSON.CreateFromParams(eJSONObject, "type", itemType)

    'Each item of the form: {id: number, data: {... updates ...}}
    Dim item As stdJSON
    For Each item In changeSet
      'Bind metadata
      With item("data")
        if not .exists("__metadata") then Call .Add("__metadata", metadata)
      End With

      'Add change
      Call batch.Push("--changeset_" & changeSetID)
      Call batch.Push("Content-Type:application/http")
      Call batch.Push("Content-Transfer-Encoding: binary")
      Call batch.Push("")
      Call batch.Push(InterpolateString("PATCH ${site}/${list}/items(${itemID}) HTTP/1.1", "site", this.siteURL, "list", This.listAPI, "itemID", item("id")))
      Call batch.Push("Content-Type: application/json;odata=verbose;")
      Call batch.Push("Accept: application/json")
      Call batch.Push("If-Match: *")
      Call batch.Push("X-HTTP-Method: MERGE")
      Call batch.Push("")
      Call batch.Push(item("data").ToString())
      Call batch.Push("")
    Next

    'End changeset to create Data
    Call batch.Push("--changeset_" & changeSetID & "--")
    Dim body As String: body = batch.Join(vbCrLf)

    'Make request via RawRequest
    Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Content-Type", "multipart/mixed; boundary=batch_" & batchUUID)
    Dim request As stdHTTP: Set request = RawRequest("_api/$batch", "POST", headers, body)
    Call responses.add(request)
  Next
  
  
  Set ListItemsUpdateBatch = responses
  Set ListItemsUpdateBatch.responseMapper = stdCallback.CreateFromObjectMethod(Me, "protProcessBatchResponse").bind("update", additions)
End Function

'Uses $batch query to delete many items in a sharepoint list in as few requests as possible
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsDeleteBatch(ByVal deletionIDs As stdArray, Optional ByVal changeSetID As String, Optional ByVal batchUUID As String) As HTTPCollection
  If changeSetID = "" Then changeSetID = getGUID()
  If batchUUID = "" Then batchUUID = getGUID()
  Dim responses As HTTPCollection: Set responses = HTTPCollection.Create()

  'A max of 1000 operations are allowed in a changeset; To stay well under this value, we use batches of 750.
  For Each changeSet In deletionIDs.SplitToChunks(750)
    'Workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
    Call changeSet.Push(changeSet.item(1))

    'Create changeset
    Dim batch As stdArray: Set batch = stdArray.Create()
    Call batch.Push("--batch_" & batchUUID)
    Call batch.Push("Content-Type: multipart/mixed; boundary=changeset_" & changeSetID)
    Call batch.Push("")

    'Each item of the form: {id: number, data: {... updates ...}}
    Dim id
    For Each id In changeSet
      'Add change
      Call batch.Push("--changeset_" & changeSetID)
      Call batch.Push("Content-Type:application/http")
      Call batch.Push("Content-Transfer-Encoding: binary")
      Call batch.Push("")
      Call batch.Push(InterpolateString("DELETE ${site}/${list}/items(${itemID}) HTTP/1.1", "site", this.siteURL, "list", This.listAPI, "itemID", id))
      Call batch.Push("Accept: application/json;odata=verbose")
      Call batch.Push("Content-Type: application/json;odata=verbose")
      Call batch.Push("If-Match: *")
      Call batch.Push("")
    Next

    'End changeset to create Data
    Call batch.Push("--changeset_" & changeSetID & "--")
    Dim body As String: body = batch.Join(vbCrLf)

    'Make request via RawRequest
    Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Content-Type", "multipart/mixed; boundary=batch_" & batchUUID)
    Dim request As stdHTTP: Set request = RawRequest("_api/$batch", "POST", headers, body)
    Call responses.add(request)
  Next
  
  Set ListItemsDeleteBatch = responses
  Set ListItemsDeleteBatch.responseMapper = stdCallback.CreateFromObjectMethod(Me, "protProcessBatchResponse").bind("delete", additions)
End Function

'@devNote ```HTTP
'POST https://{site_url}/_api/web/GetFileByServerRelativeUrl('/Shared Documents/{file_name}')/CheckOut()
'Authorization: Bearer {access_token}
'X-RequestDigest: {form_digest_value}
'```
Public Function DocLibraryFileCheckOut(ByVal relativePath As String) As stdHTTP
  Const APITemplate As String = "_api/web/GetFileByServerRelativeUrl('/${docLib}/${relativePath}')/CheckOut()"
  Dim API As String: API = InterpolateString(APITemplate, "docLib", This.listSelectorValue, "relativePath", relativePath)
  'TODO: X-RequestDigest?
  Dim headers As stdJSON: Set headers = stdJSON.Create()

  Set DocLibraryFileCheckOut = RawRequest(API, "POST", headers)
End Function

'@devNote ```HTTP
'POST https://{site_url}/_api/web/GetFileByServerRelativeUrl('/Shared Documents/{file_name}')/CheckIn(comment='Comment',checkintype=0)
'Authorization: Bearer {access_token}
'X-RequestDigest: {form_digest_value}
'```
Public Function DocLibraryFileCheckIn(ByVal relativePath As String, ByVal versionComment As String, ByVal iType As SharepointFileCheckInType)
  Const APITemplate As String = "_api/web/GetFileByServerRelativeUrl('/${docLib}/${relativePath}')/CheckIn(comment='${comment}', checkintype=${checkInType})"
  Dim API As String: API = InterpolateString(APITemplate, "docLib", This.listSelectorValue, "relativePath", relativePath, "comment", versionComment, "checkInType", iType)
  'TODO: X-RequestDigest?
  Dim headers As stdJSON: Set headers = stdJSON.Create()

  Set DocLibraryFileCheckIn = RawRequest(API, "POST", headers)
End Function

'@devNote ```HTTP
'GET https://{site_url}/_api/web/GetFolderByServerRelativeUrl('Folder Name')/Files
'Authorization: Bearer {access_token}
'Accept: application/json;odata=verbose
'```
Public Function DocLibraryFilesGet()
  'TODO:
End Function

'@devNote ```HTTP
'GET https://{site_url}/_api/web/GetFileByServerRelativeUrl('/Folder Name/{file_name}')/$value
'Authorization: Bearer {access_token}
'```
Public Function DocLibraryFileGet()
  'TODO:
End Function

'@devNote ```HTTP
'POST https://{site_url}/_api/web/GetFileByServerRelativeUrl('/Folder Name/{file_name}')/$value
'Authorization: Bearer {access_token}
'Content-Length: {length}
'X-HTTP-Method: PUT
'X-RequestDigest: {form_digest_value}
'
'Contents of file
'```
Public Function DocLibraryFileUpdate()

End Function

'@devNote ```HTTP
'POST https://{site_url}/_api/web/GetFolderByServerRelativeUrl('Folder Name')/Files/add(url='a.txt',overwrite=true)
'Authorization: Bearer {access_token}
'Content-Length: {length}
'X-RequestDigest: {form_digest_value}
'
'Contents of file
'```
Public Function DocLibraryFileCreate()

End Function



'@param listSelector - title/id/name of the sharepoint list
'@param ListSelectorType - SharepointListByTitle (default)/SharepointListByID/SharepointListByName respectively compared to `listSelector`
'@returns - HTTP response. Use `http.ResponseMapped()` to get the ItemType
Public Function getListItemType()
  Const APITemplate As String = "${list}?$select=ListItemEntityTypeFullName"
  Dim API As String: API = InterpolateString(APITemplate, "list", This.listAPI)
  Dim headers As stdJSON: Set headers = stdJSON.CreateFromParams(eJSONObject, "Accept", "application/json;odata=verbose")

  Set getListItemType = RawRequest(API, "GET", headers)
  Set getListItemType.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText).item(""d"").item(""ListItemEntityTypeFullName"")").Bind(stdJSON)
End Function




'Get a GUID, used in $batch queries
'@returns - A new GUID
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function

'Interpolate a string
'@param input - template string
'@param replacements - replacements
'@returns - The interpolated string
Private Function InterpolateString(ByVal sInput As String, ParamArray replacements()) As String
  Dim i As Long
  For i = 0 To UBound(replacements) Step 2
    If InStr(1, sInput, "${" & replacements(i) & "}") > 0 Then
      sInput = Replace(sInput, "${" & replacements(i) & "}", replacements(i + 1))
    End If
  Next
  InterpolateString = sInput
End Function
