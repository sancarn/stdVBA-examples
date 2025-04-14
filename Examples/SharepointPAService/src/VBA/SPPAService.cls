VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SPPAService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TThis
  serviceWorker as string
  siteURL as string
  listSelectorType as string
  listSelectorValue as string
  listAPI as string
End Type
Private This as TThis

'Create a power automate service worker to automate the creation, maintainence and deletion of sharepoint items
'@constructor
'@param serviceWorker - The url to the power automate service worker.
'@param siteURL - The sharepoint site url.
'@param listSelectorValue - title, id or name of the list to run the queries against.
'@param listSelectorType - "title", "id" or "name" depending on how you are wanting to select the list in question.
'@returns - Sharepoint PowerAutomate service worker.
'@example `set service = SPPAService.Create("http://service.worker.com", "http://tenant.sharepoint.com/sites/My/Site", "My List")`
Public Function Create(ByVal serviceWorker as string, ByVal siteURL as string, listSelectorValue as string, optional ByVal listSelectorType as string = "title") as SPPAService
  set Create = new SPPAService
  Call Create.protInit(serviceWorker, siteURL, listSelectorValue, listSelectorType)
End Function

'Initialisation sub
'@protected
'@param siteURL - The sharepoint site url
Public Sub protInit(serviceWorker as string, siteURL as string, listSelectorValue as string, listSelectorType as string)
  This.serviceWorker = serviceWorker
  This.siteURL = siteURL
  This.listSelectorType = listSelectorType
  This.listSelectorValue = listSelectorValue
  
  Const APITemplate as string = "_api/web/lists${selector}"
  select case lcase(listSelectorType)
    case "title"
      This.listAPI = InterpolateString(APITemplate, "selector", "/GetByTitle('${title}')", "title", listSelectorValue)
    case "id"
      This.listAPI = InterpolateString(APITemplate, "selector", "guid('${guid}')", "guid", listSelectorValue)
    case "name"
      This.listAPI = InterpolateString(APITemplate, "selector", "/" & listSelectorValue)
  end select
End Sub


'Create a power automate service worker from a configuration file. These configuration files can be created with the `ToConfigFile` method.
'@constructor
'@param workerFile - Path to a worker config file. The file should be in the JSON format with a `url` property.
'@param listOrLibraryFile - Path to a list/library config file. The file should be in the JSON format with a `siteURL`, `selectorType` and `selectorValue` properties.
'@returns - Sharepoint PowerAutomate service worker.
'@example `set service = SPPAService.CreateFromConfigs("C:\Tmp\MyWorker.json", "C:\Tmp\MyList.json")`
Public Function CreateFromConfigs(ByVal workerFile as string, ByVal listOrLibraryFile as string) as SPPAService
  Dim workerConfig as stdJSON: set workerConfig = stdJSON.CreateFromFile(workerFile)
  Dim llConfig as stdJSON: set llConfig = stdJSON.CreateFromFile(listOrLibraryFile)
  set CreateFromConfigs = Create(workerConfig("url"), llConfig("siteURL"), llConfig("selectorValue"), llConfig("selectorType"))
End Function


'Create a power automate service worker from a configuration file. These configuration files can be created with the `ToConfigFile` method.
'@constructor
'@param configFile - Path to a config file. The file should be in the JSON format with a `serviceWorker` and a `siteURL` property.
'@returns - Sharepoint PowerAutomate service worker.
'@example ```vba
'SPPAService.CreateFromCombinedConfig("C:\Tmp\MyConnectionFile.json")
'```
Public Function CreateFromCombinedConfig(ByVal configFile as string) as SPPAService
  Dim config as stdJSON: set config = stdJSON.CreateFromFile(configFile)
  set CreateFromCombinedConfig = Create(config("serviceWorker"), config("list")("siteURL"), config("list")("selectorValue"), config("list")("selectorType"))
End Function


'Creates a worker config file from the existing instance
'@param configFile - Path to the desired config file location
Public Sub ToWorkerConfig(ByVal configFile as string)
  With stdJSON.Create()
    .add "url", this.serviceWorker
    Call .ToFile(configFile,2)
  End With
End Sub

'Creates a list/library config file from the existing instance
'@param configFile - Path to the desired config file location
Public Sub ToListLibConfig(ByVal configFile as string)
  With stdJSON.Create()
    .add "siteURL", this.siteURL
    .add "selectorValue", this.listSelectorValue
    .add "selectorType", this.listSelectorType
    Call .ToFile(configFile,2)
  End With
End Sub

'Creates a config file from the existing instance
'@param configFile - Path to the config file to be produced
Public Sub ToCombinedConfig(ByVal configFile as string)
  With stdJSON.Create()
    .add "serviceWorker", this.serviceWorker
    With .addObject("list")
      .add "siteURL", this.siteURL
      .add "selectorValue", this.listSelectorValue
      .add "selectorType", this.listSelectorType
    End With
    Call .ToFile(configFile,2)
  End With
End Sub

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
  stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true) _
)
'Debug.Print http.ResponseText
'```
'@remarks All responses have a `ResponseMapper` to parse the response text as JSON. This may need to be overriden however.
Public Function RawRequest(ByVal apiPath as string, ByVal method as string, ByVal headers as stdJSON, Optional ByVal body as string = "") as stdHTTP
  With stdJSON.Create()
    .add "site", this.serviceWorker
    .add "api", apiPath
    .add "method", method
    .add "headers", headers
    .add "body", body

    Dim body as string: body = .toString()

    set RawRequest = stdHTTP.Create( _
      This.serviceWorker, _
      ThreadingStyle:= HTTPAsync, _
      options:= stdHTTP.CreateOptions(body) _
    )

    'Always assume JSON response, but allow override
    set RawRequest.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText)").bind(stdJSON)
  End With
End Function

'Get a single item from a list by it's id
'@param itemID - the item to return
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemGet(ByVal itemID as long) as stdHTTP
  Const APITemplate as string = "${list}/items(${itemID})"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI, "itemID", itemID)
  Dim headers as stdJSON: set headers = stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true)

  set ListItemGet = RawRequest(API, "GET", headers)
End Function

'Create a single item in a list
'@param data - The data the new item should contain
'@param itemType - Optional. If not supplied it will be queried at runtime on every call. Query once to increase performance, or add as a known value. Can be obtained using `getListItemType` function
'@param ListSelectorType - SharepointListByTitle (default)/SharepointListByID/SharepointListByName respectively compared to `listSelector`
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemCreate(ByVal data as stdJSON, Optional ByVal itemType as string = "") as stdHTTP
  'Prepare item type
  if itemType = "" then itemType = getListItemType().Await().ResponseMapped()
  With data.addObject("__metadata")
    .add "type", itemType
  End With

  Const APITemplate as string = "${list}/items"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI)
  Dim headers as stdJSON
  With stdJSON.Create()
    .add "Accept", "application/json;odata=verbose"
    .add "Content-Type", "application/json;odata=verbose"
    set headers = .self
  End with

  set ListItemCreate = RawRequest(API, "POST", headers, data.toString())
End Sub

'Update a single item in a list
'@param itemID - Long representing the item to update's internal ID
'@param data - The data the new item should contain
'@param itemType - Optional. If not supplied it will be queried at runtime on every call. Query once to increase performance, or add as a known value. Can be obtained using `getListItemType` function
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemUpdate(ByVal itemID as Long, ByVal data as stdJSON, Optional ByVal itemType as string = "") as stdHTTP
  'Prepare item type
  if itemType = "" then itemType = getListItemType().Await().ResponseMapped()
  With data.addObject("__metadata")
    .add "type", itemType
  End With

  Const APITemplate as string = "${list}/items(${itemID})"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI, "itemID", itemID)
  Dim headers as stdJSON
  With stdJSON.Create()
    .add "Accept", "application/json;odata=verbose"
    .add "Content-Type", "application/json;odata=verbose"
    .add "If-Match", "*"
    .add "X-HTTP-Method", "MERGE"
    set headers = .self
  End With

  set ListItemUpdate = RawRequest(API, "POST", headers, data.toString())
End Function

'Delete a single item in a list
'@param itemID - Long representing the item to update's internal ID
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItemDelete(ByVal itemID as Long) as stdHTTP
  Const APITemplate as string = "${list}/items(${itemID})"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI, "itemID", itemID)
  Dim headers as stdJSON
  With stdJSON.Create()
    .add "Accept", "application/json;odata=verbose"
    .add "Content-Type", "application/json;odata=verbose"
    .add "If-Match", "*"
    .add "X-HTTP-Method", "DELETE"
    set headers = .self
  End With

  set ListItemDelete = RawRequest(API, "POST", headers)
End Function

'Get all list items from a sharepoint list
'@returns - HTTP response. Use `http.ResponseMapped()` to get returned data
Public Function ListItems() as stdHTTP
  Const APITemplate as string = "${list}/items"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI)
  Dim headers as stdJSON: set headers = stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true)

  set ListItems = RawRequest(API, "GET", headers)
End Function

'Get list items filtering to an ODATA Query
'@example `myList.ListItemsFromODataQuery("myCol eq 'test'")`
'@remark ODATA Queries do not work effectively on lists with more than 500 records. It's recommended to use `ListItemsFromInplaceQuery` instead where possible.
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
Public Function ListItemsFromODataQuery(ByVal query as string) as stdJSON
  Const APITemplate as string = "${list}/items?$filter=${oDataQuery}"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI, "oDataQuery", query)
  Dim headers as stdJSON: set headers = stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true)
  Dim data as stdHTTP: set data = RawRequest(API, "GET", headers).Await().ResponseMapped().item("d")
  Dim result as stdJSON: set result = stdJSON.Create(eJSONArray)
  Call result.add(data.item("results"))

  While not isNull(data.item("__next"))
    set data = RawRequest(data.item("__next"), "GET", headers).Await().ResponseMapped().item("d")
    Call result.add(data.item("results"))
  Wend

  set ListItemsFromODataQuery = result
End Function

'Get all items using an inplace query
Public Function ListItemsFromInplaceQuery(ByVal query as string) as stdHTTP
  Const APITemplate as string = "${list}/RenderListDataAsStream?InplaceSearchQuery=${inplaceQuery}"
  Dim API as string: API = InterpolateString(API, "list", this.listAPI, "inplaceQuery", query)
  Dim headers as stdJSON: set headers = stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true)

  set ListItemsFromInplaceQuery = RawRequest(API, "POST", headers)
  set ListItemsFromInplaceQuery.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText).Row").bind(stdJSON) 
End Function

'Uses $batch query to add many items to a sharepoint list in as few requests as possible
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsCreate(ByVal additions as stdEnumerator, Optional ByVal itemType as string = "") as stdHTTP
  'TODO: 
End Function

'Uses $batch query to update many items in a sharepoint list in as few requests as possible
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsUpdate(ByVal changes as stdEnumerator, Optional ByVal itemType as string = "") as stdHTTP
  'TODO: 
End Function

'Uses $batch query to delete many items in a sharepoint list in as few requests as possible
'@docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
Public Function ListItemsDelete(ByVal deletionIDs as stdEnumerator) as stdHTTP
  'TODO: 
End Function

'@param listSelector - title/id/name of the sharepoint list
'@param ListSelectorType - SharepointListByTitle (default)/SharepointListByID/SharepointListByName respectively compared to `listSelector`
'@returns - HTTP response. Use `http.ResponseMapped()` to get the ItemType
Public Function getListItemType()
  Const APITemplate as string = "${list}?$select=ListItemEntityTypeFullName"
  Dim API as string: API = InterpolateString(APITemplate, "list", this.listAPI)
  Dim headers as stdJSON: set headers = stdJSON.CreateFromString("{Accept: ""application/json;odata=verbose""}", true)

  set getListItemType = RawRequest(API, "GET", headers)
  set getListItemType.ResponseMapper = stdLambda.Create("$1.CreateFromString($2.ResponseText).d.ListItemEntityTypeFullName").bind(stdJSON)
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
Private Function InterpolateString(ByVal input as string, ParamArray replacements()) as string
  Dim i as long
  For i = 0 to ubound(replacements) step 2
    if instr(1, input, "${" & replacements(i) & "}") > 0 then
      input = replace(input, "${" & replacements(i) & "}", replacements(i+1))
    end if
  next
  InterpolateString = input
End Function
