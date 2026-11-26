# SharePoint Updator (VBA)

This project is primarily a reusable SharePoint integration toolkit for VBA, centered on:
- `src/stdSharepointAuthenticator.frm` for browser-session authentication and cookie injection.
- `src/stdSharepointList.cls` for SharePoint list CRUD, querying, transforms, and batching.

Examples:
- `examples/ORMMissingData/*` — typed wrapper for a specific business list.
- `examples/GenericUpdater/*` — range-based generic batch updater driven by SiteURL + ListTitle.

## Core components

- `src/stdSharepointAuthenticator.frm`  
  Implements `stdICallable` so it can be passed into `stdHTTP` requests as an authenticator. It lazily opens a WebView, authenticates against SharePoint, and attaches the correct `Cookie` header for subsequent requests.

- `src/stdSharepointList.cls`  
  Generic SharePoint list REST client built for VBA. It handles list URL parsing, OData query generation, item transforms, form digest management, and batch request/response handling.

## What this enables

Using `stdSharepointAuthenticator` + `stdSharepointList`, you can:
- authenticate once through a real SharePoint web session,
- read single items and paged item collections,
- run in-place list queries for large lists,
- create, update, and delete items,
- execute high-volume batch operations,
- map SharePoint field types (person, multiperson, multichoice, lookup, date) to/from VBA-friendly payloads.

## Quick usage

```vb
Dim auth As stdSharepointAuthenticator
Set auth = stdSharepointAuthenticator.Create("https://contoso.sharepoint.com")

Dim list As stdSharepointList
Set list = stdSharepointList.Create( _
  "https://contoso.sharepoint.com/sites/Projects/Lists/Risks/AllItems.aspx", _
  auth _
)

Call list.FieldsAdd("Title", SharePointFieldText)
Call list.FieldsAdd("Owner", SharePointFieldPerson)

Dim rows As stdJSON
Set rows = list.ItemsGet()
Debug.Print rows.Length
```

You can also construct a list client from site URL + display title:

```vb
Set list = stdSharepointList.CreateFromTitle( _
  "https://contoso.sharepoint.com/sites/Projects", _
  "Risks", _
  auth _
)
Dim schema As stdJSON
Set schema = list.FieldsFetchSchema()
```

## Example folders

### ORMMissingData (typed wrapper)

- `examples/ORMMissingData/ORMMissingDataRow.cls`
- `examples/ORMMissingData/ContosoAuth.bas`
- `examples/ORMMissingData/mMain.bas`

Sample domain code that wraps `stdSharepointList` for a specific business list. Not the main library surface.

### GenericUpdater (range-based batch updater)

- [`examples/GenericUpdater/README.md`](examples/GenericUpdater/README.md) — full usage, cell encoding, and field-type special cases (person email, multi-value splits, `null` clears, etc.)
- `examples/GenericUpdater/src/GenericUpdater.bas`
- `examples/GenericUpdater/src/mMain.bas`

Range-driven Create/Update for arbitrary lists (`UpdateList`, `DumpResults`, `EnsureTableFields`). Demo macros: `mainRunGenericUpdater`, `mainFetchFields`.

## Requirements

- stdVBA dependencies used by this project (including `stdHTTP`, `stdJSON`, `stdICallable`, `stdWebView`, and the bundled `stdOutlook` class for optional sign-in prefill).
- A SharePoint tenant and permissions for the target list/site.
- Correct site/list URLs and field internal names in your consuming code.
