# xlSaveHandler

Intercepts Excel's "Do you want to save your workbook" message, and raises various events allowing more control over the workbook.

Events raised:

* `BeforeShow(obj As xlSaveHandler)` - Before the UI shows
* `AfterShow(obj As xlSaveHandler)` - After the UI shows
* `WorkbookCancelSave()` - After the UI shows, if the user clicked cancel
* `WorkbookBeforeSave()` - After the UI shows, before save is called
* `WorkbookAfterSave()` - After the UI shows, after save is called
* `WorkbookClose()` - Immediately before the workbook is closed.