Attribute VB_Name = "mMain"

'@module
'@description Generic updater workbook entry points.
'@aidate 2026-11-26
'@aidesc AI added output sheet clearing at the start of each upload.
Option Explicit

'Demo entry point for GenericUpdater.
'Point PatchData at a headered table with ID, optional Type (Create|Update), and field columns.
'Blank cells skip the field; type the word null to clear it.
Public Sub mainRunGenericUpdater()
  shOutput.Cells.Clear
  Dim SiteURL As String: SiteURL = shInput.Range("SITE_URL").Text
  Dim ListTitle As String: ListTitle = shInput.Range("LIST_NAME").Text
  Dim PatchData As Range: Set PatchData = shInput.ListObjects("DataToUpdate").Range
  Dim Dest As Range: Set Dest = shOutput.Range("A1")

  Dim results As Variant: results = GenericUpdater.UpdateList(SiteURL, ListTitle, PatchData)
  Call GenericUpdater.DumpResults(results, Dest)
  MsgBox "Jobs done"
End Sub

'Fetch writable SharePoint fields and add any missing columns to DataToUpdate (keeps existing data).
Public Sub mainFetchFields()
  Dim SiteURL As String: SiteURL = shInput.Range("SITE_URL").Text
  Dim ListTitle As String: ListTitle = shInput.Range("LIST_NAME").Text
  Dim Table As ListObject: Set Table = shInput.ListObjects("DataToUpdate")

  Dim added As Long: added = GenericUpdater.EnsureTableFields(SiteURL, ListTitle, Table)
  MsgBox "Added " & CStr(added) & " column(s) to DataToUpdate."
End Sub
