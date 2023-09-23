# xlApplication

This class is `Work In Progress`.

Initially it's main purpose will be to help obtain various information about the Excel application which is difficult to obtain otherwise.

Long term the hope is to create a layer over the existing Excel API which is stdVBA complient.

* `CreateFromApplication` - Creates from an existing instance
* `CreateFromHWND` - Creates from an existing instance via HWND
* `CreateFromAllApplications` - Creates a collection from all existing Excel application instances
* `WIP CreateNewInstance(Optional AccesVBOM = vbYes)` - Create a new instance of Application. Optionally enable VBOM, default yes.
* `WIP Get EditMode`: will return `Undefined`, `Ready`, `Enter`, `Edit`, `Point` dependent on whether the 
* `WIP Get VBARuntimeType`: will return `Running`, `Break` or `Stopped`. This ultimately boils down to whether the code is running or not, and can be found ordinarily in the caption of the VBA window, but can be read even while VBA is running (i.e. by another application)
* `WIP Get VBAMode` - a.k.a `EbMode`. Can currently only be run from the instance itself. Returns `true` if the code is 'running' and `false` if not. I.E. This will be true if any code has run, global objects are alive etc. It will only switch back to false when clicking the Stop button.
* `WIP Get VBOM` - returns `stdEnumerator<Workbook>`
* `WIP Get Workbooks` - returns `stdEnumerator<Object<xlWorkbook>>`
* ...