Attribute VB_Name = "GenericUpdater"

'@module
'@aidate 2026-11-26
'@aidesc AI added timeout handling and batched person resolution with row-level errors.
'@description Range-based generic SharePoint list updater / creator.
'Call order:
'  1. stdSharepointAuthenticator.Create(SiteURL)
'  2. auth.protEnsureAuthenticated
'  3. stdSharepointList.CreateFromTitle(SiteURL, ListTitle, auth)
'  4. list.FieldsFetchSchema
'  5. VerifyThatFieldNamesExist (fail all rows if mismatch, no write)
'  6. Verify individual data entries against list field rules
'  7. BatchItemsCreate / BatchItemsSet in 500-item chunks by Type column
'
'Reserved columns:
'  - ID    - required for Update rows; ignored for Create (new Id written to results on success)
'  - Type  - Create | Update (if omitted, all rows are treated as Update)
'
'Cell encoding:
'  - Empty / blank -> skip (omit from payload; leave existing value on Update)
'  - "null" (after trim, case-insensitive) -> clear (send JSON null)
'  - Any other value -> validate, then include
'@example ```vb
'Dim results As Variant: results = GenericUpdater.UpdateList( _
'  "https://contoso.sharepoint.com/sites/Projects", _
'  "Risks", _
'  Sheet1.Range("A1").CurrentRegion _
')
'Call GenericUpdater.DumpResults(results, Sheet1.Range("F1"))
'```

Option Explicit

Private Const BATCH_SIZE As Long = 500
Private Const RECOVERY_KEY As String = "__create_recovery_key__"
Private Const CLEAR_TOKEN As String = "null"

Private Enum ERowOp
  RowOpUpdate = 1
  RowOpCreate = 2
End Enum

Private Type TFieldSchema
  Title As String
  InternalName As String
  TypeAsString As String
  FieldType As SharePointFieldType
  Required As Boolean
  ReadOnlyField As Boolean
  Hidden As Boolean
  MaxLength As Long
  HasMaxLength As Boolean
  choices As Object ' Scripting.Dictionary, key=choice text (vbTextCompare)
  AllowMultipleValues As Boolean
  Writable As Boolean
End Type

Private Type TColumnMap
  header As String
  ColIndex As Long
  schemaIndex As Long
  InternalName As String
  FieldType As SharePointFieldType
  Required As Boolean
  HasMaxLength As Boolean
  MaxLength As Long
  choices As Object
End Type

'Create / update SharePoint list items from a headered Excel range.
'@param SiteURL - Absolute SharePoint site URL.
'@param ListTitle - List display title.
'@param PatchData - Range including headers. Must contain an ID column. Optional Type column (Create|Update).
'@returns Variant 1-based 2D array with columns ID | SuccessFailure | FailureReason.
Public Function UpdateList( _
  ByVal SiteURL As String, _
  ByVal ListTitle As String, _
  ByVal PatchData As Range _
) As Variant
  If PatchData Is Nothing Then
    Err.Raise 5, "GenericUpdater::UpdateList", "PatchData cannot be Nothing."
  End If
  If PatchData.rows.count < 2 Or PatchData.columns.count < 1 Then
    Err.Raise 5, "GenericUpdater::UpdateList", "PatchData must include a header row and at least one data row."
  End If

  Dim patchValues As Variant: patchValues = PatchData.Value

  Dim idCol As Long: idCol = FindNamedColumn(patchValues, "ID", "Id")
  If idCol = 0 Then
    Err.Raise 5, "GenericUpdater::UpdateList", "PatchData must contain a column named 'ID'."
  End If
  Dim typeCol As Long: typeCol = FindNamedColumn(patchValues, "Type")

  Dim dataRows As Long: dataRows = UBound(patchValues, 1) - 1
  Dim results As Variant: results = BuildEmptyResults(patchValues, idCol, dataRows)

  '1-2. Authenticate
  Dim auth As stdSharepointAuthenticator: Set auth = stdSharepointAuthenticator.Create(SiteURL)
  Call auth.protEnsureAuthenticated

  '3-4. List client + schema
  Dim list As stdSharepointList: Set list = stdSharepointList.CreateFromTitle(SiteURL, ListTitle, auth)
  Dim schemaJson As stdJSON: Set schemaJson = list.FieldsFetchSchema()

  Dim fields() As TFieldSchema
  Dim fieldCount As Long: fieldCount = 0
  Dim recovery As stdJSON
  Call ParseSchema(schemaJson, fields, fieldCount, recovery)

  '5. Verify field names exist (writable Title/InternalName match). ID and Type are reserved.
  Dim unmatched As Collection: Set unmatched = New Collection
  Dim columns() As TColumnMap
  Dim colCount As Long: colCount = 0
  Call MapColumns(patchValues, idCol, typeCol, fields, fieldCount, columns, colCount, unmatched, recovery)

  If unmatched.count > 0 Then
    Dim failReason As String: failReason = BuildUnmatchedReason(unmatched, ListTitle)
    Dim i As Long
    For i = 1 To dataRows
      results(i + 1, 2) = "Failure"
      results(i + 1, 3) = failReason
    Next
    UpdateList = results
    Exit Function
  End If

  'Register mapped fields for setter transforms
  Dim fieldIdx As Long
  For fieldIdx = 1 To colCount
    Call list.FieldsAdd(columns(fieldIdx).InternalName, columns(fieldIdx).FieldType, Not columns(fieldIdx).Required)
  Next

  '6. Validate individual rows; build create/update payloads for valid rows
  Dim updatePayloads As stdJSON: Set updatePayloads = stdJSON.Create(eJSONArray)
  Dim createPayloads As stdJSON: Set createPayloads = stdJSON.Create(eJSONArray)
  Dim validRowIndexes As Collection: Set validRowIndexes = New Collection
  Dim updateRowById As Object: Set updateRowById = CreateObject("Scripting.Dictionary")
  updateRowById.compareMode = vbTextCompare
  Dim createRowIndexes As Collection: Set createRowIndexes = New Collection

  Dim people As Object: Set people = CreateObject("Scripting.Dictionary")
  people.CompareMode = vbTextCompare
  Dim nextRow As Long: nextRow = 1
  Dim rowIndex As Long
  Do While nextRow <= dataRows
    Dim candidates As Collection: Set candidates = New Collection
    Dim pendingIds As Object: Set pendingIds = CreateObject("Scripting.Dictionary")
    Dim logins As Collection: Set logins = New Collection
    Do While nextRow <= dataRows
      Dim rowIssues As Collection: Set rowIssues = New Collection
      Dim itemId As Long: itemId = 0
      Dim rowOp As ERowOp: rowOp = RowOpUpdate
      ResolveRowOp patchValues, nextRow + 1, typeCol, rowOp, rowIssues
      If rowIssues.count = 0 Then
        ValidateRow patchValues, nextRow + 1, idCol, rowOp, columns, colCount, updateRowById, itemId, rowIssues
      End If
      If rowIssues.count = 0 Then
        If rowOp = RowOpUpdate Then
          'Resolve the earlier row before deciding whether this duplicate may be submitted.
          If pendingIds.Exists(itemId) Then Exit Do
          pendingIds.Add itemId, True
        End If
        CollectRowLogins patchValues, nextRow + 1, columns, colCount, people, logins
        candidates.Add Array(nextRow, rowOp, itemId)
      Else
        results(nextRow + 1, 2) = "Failure"
        results(nextRow + 1, 3) = JoinCollection(rowIssues, "; ")
      End If
      nextRow = nextRow + 1
    Loop

    Dim retryable As Collection: Set retryable = New Collection
    list.BatchEnsureUsers logins, people, retryable
    Dim candidate As Variant
    For Each candidate In candidates
      rowIndex = candidate(0)
      rowOp = candidate(1)
      itemId = candidate(2)
      Set rowIssues = New Collection
      Dim dataObj As stdJSON: Set dataObj = BuildRowPayload(patchValues, rowIndex + 1, columns, colCount, people, rowIssues)
      If rowIssues.count > 0 Then
        results(rowIndex + 1, 2) = "Failure"
        results(rowIndex + 1, 3) = JoinCollection(rowIssues, "; ")
      Else
        validRowIndexes.Add rowIndex
        If rowOp = RowOpCreate Then
          createPayloads.Add dataObj
          createRowIndexes.Add rowIndex
        Else
          Dim payload As stdJSON: Set payload = stdJSON.Create(eJSONObject)
          payload.Add "id", itemId
          payload.Add "data", dataObj
          updatePayloads.Add payload
          updateRowById.Add CStr(itemId), rowIndex
        End If
      End If
    Next
    Dim login As Variant
    For Each login In retryable
      people.Remove login
    Next
  Loop

  If updatePayloads.Length = 0 And createPayloads.Length = 0 Then
    UpdateList = results
    Exit Function
  End If

  '7. Batch writes
  Dim totalBatches As Long
  totalBatches = BatchCount(updatePayloads.Length) + BatchCount(createPayloads.Length)
  Dim processedBatches As Long: processedBatches = 0

  If updatePayloads.Length > 0 Then
    Call RunUpdateBatches(list, updatePayloads, updateRowById, results, processedBatches, totalBatches)
  End If
  If createPayloads.Length > 0 Then
    Dim recoveryProperty As String
    If Not recovery Is Nothing Then
      If LCase$(NzString(recovery, "TypeAsString")) = "text" And Not NzBool(recovery, "ReadOnlyField") Then
        If IsNull(recovery("MaxLength")) Or Val(NzString(recovery, "MaxLength")) >= 36 Then recoveryProperty = NzString(recovery, "EntityPropertyName")
      End If
    End If
    Call RunCreateBatches(list, createPayloads, createRowIndexes, results, processedBatches, totalBatches, recoveryProperty)
  End If

  'Any valid rows still unmarked after batching
  Dim j As Long
  For j = 1 To validRowIndexes.count
    rowIndex = CLng(validRowIndexes(j))
    If LenB(CStr(results(rowIndex + 1, 2))) = 0 Then
      results(rowIndex + 1, 2) = "Failure"
      results(rowIndex + 1, 3) = "No batch response"
    End If
  Next

  UpdateList = results
  Application.StatusBar = False
End Function

'Dump an UpdateList results array to a sheet starting at Dest.
'@param Results - 1-based 2D array from UpdateList.
'@param Dest - Top-left destination cell.
Public Sub DumpResults(ByVal results As Variant, ByVal Dest As Range)
  If Dest Is Nothing Then
    Err.Raise 5, "GenericUpdater::DumpResults", "Dest cannot be Nothing."
  End If
  If Not IsArray(results) Then
    Err.Raise 5, "GenericUpdater::DumpResults", "Results must be a 2D array."
  End If

  Dim rowCount As Long: rowCount = UBound(results, 1) - LBound(results, 1) + 1
  Dim colCount As Long: colCount = UBound(results, 2) - LBound(results, 2) + 1
  Dest.Resize(rowCount, colCount).value = results
End Sub

'Fetch writable list fields and ensure the ListObject has Type, ID, and a column for each field.
'Existing columns/data are preserved. Missing reserved/writable columns are appended (Type/ID inserted at the front when absent).
'@param SiteURL - Absolute SharePoint site URL.
'@param ListTitle - List display title.
'@param Table - Excel ListObject used as PatchData (e.g. DataToUpdate).
'@returns - Number of columns added.
Public Function EnsureTableFields( _
  ByVal SiteURL As String, _
  ByVal ListTitle As String, _
  ByVal Table As ListObject _
) As Long
  If Table Is Nothing Then
    Err.Raise 5, "GenericUpdater::EnsureTableFields", "Table cannot be Nothing."
  End If

  Dim auth As stdSharepointAuthenticator: Set auth = stdSharepointAuthenticator.Create(SiteURL)
  Call auth.protEnsureAuthenticated

  Dim list As stdSharepointList: Set list = stdSharepointList.CreateFromTitle(SiteURL, ListTitle, auth)
  Dim schemaJson As stdJSON: Set schemaJson = list.FieldsFetchSchema()

  Dim fields() As TFieldSchema
  Dim fieldCount As Long: fieldCount = 0
  Dim recovery As stdJSON
  Call ParseSchema(schemaJson, fields, fieldCount, recovery)

  Dim added As Long: added = 0

  If FindListColumn(Table, "Type") Is Nothing Then
    Call InsertListColumn(Table, "Type", 1)
    added = added + 1
  End If
  If FindListColumn(Table, "ID", "Id") Is Nothing Then
    Dim typeCol As ListColumn: Set typeCol = FindListColumn(Table, "Type")
    Dim idPos As Long: idPos = 1
    If Not typeCol Is Nothing Then idPos = typeCol.Index + 1
    Call InsertListColumn(Table, "ID", idPos)
    added = added + 1
  End If

  Dim i As Long
  For i = 1 To fieldCount
    If Not fields(i).Writable Then GoTo ContinueField
    If fields(i).Hidden Then GoTo ContinueField
    If ListHasFieldColumn(Table, fields(i)) Then GoTo ContinueField

    Dim header As String: header = FieldHeaderName(fields(i))
    If LenB(header) = 0 Then GoTo ContinueField
    If StrComp(header, "Type", vbTextCompare) = 0 Then GoTo ContinueField
    If StrComp(header, "ID", vbTextCompare) = 0 Then GoTo ContinueField
    If StrComp(header, "Id", vbTextCompare) = 0 Then GoTo ContinueField

    Call InsertListColumn(Table, header, Table.ListColumns.Count + 1)
    added = added + 1
ContinueField:
  Next

  EnsureTableFields = added
End Function

'---------------------------------------
' Batch runners
'---------------------------------------

Private Sub RunUpdateBatches( _
  ByVal list As stdSharepointList, _
  ByVal payloads As stdJSON, _
  ByVal updateRowById As Object, _
  ByRef results As Variant, _
  ByRef processedBatches As Long, _
  ByVal totalBatches As Long _
)
  Dim chunkStart As Long: chunkStart = 1
  Do While chunkStart <= payloads.Length
    processedBatches = processedBatches + 1
    Application.StatusBar = "Processing Batches " & CStr(processedBatches) & " / " & CStr(totalBatches)
    DoEvents

    Dim chunkEnd As Long: chunkEnd = chunkStart + BATCH_SIZE - 1
    If chunkEnd > payloads.Length Then chunkEnd = payloads.Length

    Dim chunk As stdJSON: Set chunk = stdJSON.Create(eJSONArray)
    Dim j As Long
    For j = chunkStart To chunkEnd
      chunk.Add payloads(j)
    Next

    Dim errText As String: errText = vbNullString
    Dim batchOut As stdJSON: Set batchOut = Nothing
    On Error Resume Next
    Set batchOut = list.BatchItemsSet(chunk, VerifyTimeouts:=True)
    If Err.Number <> 0 Then
      errText = Err.Description
      If LenB(errText) = 0 Then errText = "Batch update request failed."
      Err.Clear
      Set batchOut = Nothing
    End If
    On Error GoTo 0

    If batchOut Is Nothing Then
      Dim failIdx As Long
      For failIdx = chunkStart To chunkEnd
        Dim idKey As String: idKey = CStr(payloads(failIdx)("id"))
        If updateRowById.Exists(idKey) Then
          Dim rowIndex As Long: rowIndex = CLng(updateRowById(idKey))
          results(rowIndex + 1, 2) = "Failure"
          results(rowIndex + 1, 3) = errText
        End If
      Next
    Else
      Dim resultIdx As Long
      For resultIdx = 1 To batchOut.Length
        Dim one As stdJSON: Set one = batchOut(resultIdx)
        idKey = CStr(one("id"))
        If updateRowById.Exists(idKey) Then
          rowIndex = CLng(updateRowById(idKey))
          Dim httpStatus As Long: httpStatus = 0
          If one.Exists("httpStatus") Then httpStatus = CLng(one("httpStatus"))
          If (httpStatus >= 200 And httpStatus < 300) Or one.Exists("verified") Then
            results(rowIndex + 1, 2) = "Success"
            results(rowIndex + 1, 3) = vbNullString
          Else
            results(rowIndex + 1, 2) = "Failure"
            If one.Exists("timeoutReason") Then
              results(rowIndex + 1, 3) = one("timeoutReason")
            Else
              results(rowIndex + 1, 3) = ExtractBatchError(one, httpStatus)
            End If
          End If
        End If
      Next
    End If

    chunkStart = chunkEnd + 1
  Loop
End Sub

Private Sub RunCreateBatches( _
  ByVal list As stdSharepointList, _
  ByVal payloads As stdJSON, _
  ByVal createRowIndexes As Collection, _
  ByRef results As Variant, _
  ByRef processedBatches As Long, _
  ByVal totalBatches As Long, _
  ByVal recoveryProperty As String _
)
  Dim chunkStart As Long: chunkStart = 1
  Do While chunkStart <= payloads.Length
    processedBatches = processedBatches + 1
    Application.StatusBar = "Processing Batches " & CStr(processedBatches) & " / " & CStr(totalBatches)
    DoEvents

    Dim chunkEnd As Long: chunkEnd = chunkStart + BATCH_SIZE - 1
    If chunkEnd > payloads.Length Then chunkEnd = payloads.Length

    Dim chunk As stdJSON: Set chunk = stdJSON.Create(eJSONArray)
    Dim j As Long
    For j = chunkStart To chunkEnd
      chunk.Add payloads(j)
    Next

    Dim errText As String: errText = vbNullString
    Dim batchOut As stdJSON: Set batchOut = Nothing
    On Error Resume Next
    Set batchOut = list.BatchItemsCreate(chunk, RecoveryProperty:=recoveryProperty)
    If Err.Number <> 0 Then
      errText = Err.Description
      If Err.source = "stdSharepointList::BatchWriteTimeout" Then errText = "Create timed out. Outcome unknown. Check for an existing item before retrying."
      If LenB(errText) = 0 Then errText = "Batch create request failed."
      Err.Clear
      Set batchOut = Nothing
    End If
    On Error GoTo 0

    If batchOut Is Nothing Then
      Dim failIdx As Long
      For failIdx = chunkStart To chunkEnd
        Dim rowIndex As Long: rowIndex = CLng(createRowIndexes(failIdx))
        results(rowIndex + 1, 2) = "Failure"
        results(rowIndex + 1, 3) = errText
      Next
    Else
      Dim resultIdx As Long
      For resultIdx = 1 To batchOut.Length
        Dim payloadIndex As Long: payloadIndex = chunkStart + resultIdx - 1
        If payloadIndex > createRowIndexes.count Then Exit For
        rowIndex = CLng(createRowIndexes(payloadIndex))
        Dim one As stdJSON: Set one = batchOut(resultIdx)
        Dim httpStatus As Long: httpStatus = 0
        If one.Exists("httpStatus") Then httpStatus = CLng(one("httpStatus"))
        If (httpStatus >= 200 And httpStatus < 300) Or one.Exists("verified") Then
          results(rowIndex + 1, 2) = "Success"
          results(rowIndex + 1, 3) = vbNullString
          If one.Exists("id") Then results(rowIndex + 1, 1) = one("id")
        Else
          results(rowIndex + 1, 2) = "Failure"
          If one.Exists("timeoutReason") Then
            results(rowIndex + 1, 3) = one("timeoutReason")
          ElseIf httpStatus = 408 Or httpStatus = 504 Then
            results(rowIndex + 1, 3) = "Create timed out. Outcome unknown. Check for an existing item before retrying."
          Else
            results(rowIndex + 1, 3) = ExtractBatchError(one, httpStatus)
          End If
        End If
      Next
    End If

    chunkStart = chunkEnd + 1
  Loop
End Sub

Private Function BatchCount(ByVal itemCount As Long) As Long
  If itemCount <= 0 Then Exit Function
  BatchCount = (itemCount + BATCH_SIZE - 1) \ BATCH_SIZE
End Function

'---------------------------------------
' Schema / column mapping
'---------------------------------------

Private Function FindNamedColumn(ByRef patchValues As Variant, ParamArray names() As Variant) As Long
  Dim c As Long
  For c = LBound(patchValues, 2) To UBound(patchValues, 2)
    Dim header As String: header = Trim$(CStr(patchValues(1, c)))
    Dim n As Long
    For n = LBound(names) To UBound(names)
      If StrComp(header, CStr(names(n)), vbTextCompare) = 0 Then
        FindNamedColumn = c
        Exit Function
      End If
    Next
  Next
End Function

Private Sub ResolveRowOp( _
  ByRef patchValues As Variant, _
  ByVal sheetRow As Long, _
  ByVal typeCol As Long, _
  ByRef rowOp As ERowOp, _
  ByVal issues As Collection _
)
  'No Type column => Update (backward compatible)
  If typeCol = 0 Then
    rowOp = RowOpUpdate
    Exit Sub
  End If

  Dim raw As Variant: raw = patchValues(sheetRow, typeCol)
  If IsError(raw) Then
    issues.Add "Type is an Excel error"
    Exit Sub
  End If
  If IsBlankCell(raw) Then
    issues.Add "Type is blank (expected Create or Update)"
    Exit Sub
  End If

  Dim textVal As String: textVal = Trim$(CStr(raw))
  Select Case LCase$(textVal)
    Case "update"
      rowOp = RowOpUpdate
    Case "create"
      rowOp = RowOpCreate
    Case Else
      issues.Add "Type must be Create or Update"
  End Select
End Sub

Private Sub ParseSchema(ByVal schemaJson As stdJSON, ByRef fields() As TFieldSchema, ByRef fieldCount As Long, ByRef recovery As stdJSON)
  fieldCount = 0
  Erase fields
  If schemaJson Is Nothing Or schemaJson.Length = 0 Then Exit Sub

  ReDim fields(1 To schemaJson.Length)
  Dim i As Long
  For i = 1 To schemaJson.Length
    If Not IsObject(schemaJson(i)) Then GoTo ContinueParse
    If Not TypeOf schemaJson(i) Is stdJSON Then GoTo ContinueParse
    Dim row As stdJSON: Set row = schemaJson(i)
    If Not row.Exists("InternalName") Then GoTo ContinueParse

    If StrComp(NzString(row, "Title"), RECOVERY_KEY, vbTextCompare) = 0 Or _
       StrComp(NzString(row, "InternalName"), RECOVERY_KEY, vbTextCompare) = 0 Then
      If Not recovery Is Nothing Then Err.Raise 5, "GenericUpdater", "Ambiguous create recovery field."
      Set recovery = row
      GoTo ContinueParse
    End If
    fieldCount = fieldCount + 1
    fields(fieldCount).Title = Trim$(CStr(NzString(row, "Title")))
    fields(fieldCount).InternalName = Trim$(CStr(NzString(row, "InternalName")))
    Dim typeName As String: typeName = LCase$(Trim$(CStr(NzString(row, "TypeAsString"))))
    fields(fieldCount).TypeAsString = typeName
    fields(fieldCount).FieldType = MapTypeAsString(typeName, row)
    fields(fieldCount).Required = CBool(NzBool(row, "Required"))
    fields(fieldCount).ReadOnlyField = CBool(NzBool(row, "ReadOnlyField"))
    fields(fieldCount).Hidden = CBool(NzBool(row, "Hidden"))
    fields(fieldCount).AllowMultipleValues = CBool(NzBool(row, "AllowMultipleValues"))
    fields(fieldCount).HasMaxLength = False
    fields(fieldCount).MaxLength = 0
    If row.Exists("MaxLength") Then
      If Not IsNull(row("MaxLength")) And IsNumeric(row("MaxLength")) Then
        fields(fieldCount).HasMaxLength = True
        fields(fieldCount).MaxLength = CLng(row("MaxLength"))
      End If
    End If

    Set fields(fieldCount).choices = CreateObject("Scripting.Dictionary")
    fields(fieldCount).choices.compareMode = vbTextCompare
    If row.Exists("Choices") And IsObject(row("Choices")) And TypeOf row("Choices") Is stdJSON Then
      Dim choices As stdJSON: Set choices = row("Choices")
      If choices.JsonType = eJSONArray Then
        Dim j As Long
        For j = 1 To choices.Length
          Dim choiceText As String: choiceText = Trim$(CStr(choices(j)))
          If LenB(choiceText) > 0 And Not fields(fieldCount).choices.Exists(choiceText) Then
            fields(fieldCount).choices.Add choiceText, True
          End If
        Next
      End If
    End If

    fields(fieldCount).Writable = IsWritableField(fields(fieldCount))
ContinueParse:
  Next

  If fieldCount = 0 Then
    Erase fields
  ElseIf fieldCount < schemaJson.Length Then
    ReDim Preserve fields(1 To fieldCount)
  End If
End Sub

Private Function MapTypeAsString(ByVal typeName As String, ByVal row As stdJSON) As SharePointFieldType
  Select Case typeName
    Case "text", "note"
      MapTypeAsString = SharePointFieldText
    Case "number", "currency", "integer", "counter"
      MapTypeAsString = SharePointFieldNumber
    Case "boolean"
      MapTypeAsString = SharePointFieldBoolean
    Case "datetime"
      MapTypeAsString = SharePointFieldDate
    Case "choice"
      MapTypeAsString = SharePointFieldChoice
    Case "multichoice"
      MapTypeAsString = SharePointFieldMultiChoice
    Case "usermulti"
      MapTypeAsString = SharePointFieldMultiPerson
    Case "user"
      If row.Exists("AllowMultipleValues") And CBool(row("AllowMultipleValues")) Then
        MapTypeAsString = SharePointFieldMultiPerson
      Else
        MapTypeAsString = SharePointFieldPerson
      End If
    Case "lookup"
      MapTypeAsString = SharePointFieldLookup
    Case "file"
      MapTypeAsString = SharePointFieldFile
    Case Else
      MapTypeAsString = SharePointFieldText
  End Select
End Function

Private Function IsWritableField(ByRef field As TFieldSchema) As Boolean
  If field.ReadOnlyField Then Exit Function
  Dim typeName As String: typeName = LCase$(field.TypeAsString)
  Select Case typeName
    Case "calculated", "computed", "attachments", "counter"
      Exit Function
  End Select
  If StrComp(field.InternalName, "Id", vbTextCompare) = 0 Then Exit Function
  If StrComp(field.InternalName, "ID", vbTextCompare) = 0 Then Exit Function
  IsWritableField = True
End Function

Private Sub MapColumns( _
  ByRef patchValues As Variant, _
  ByVal idCol As Long, _
  ByVal typeCol As Long, _
  ByRef fields() As TFieldSchema, _
  ByVal fieldCount As Long, _
  ByRef columns() As TColumnMap, _
  ByRef colCount As Long, _
  ByVal unmatched As Collection, _
  ByVal recovery As stdJSON _
)
  colCount = 0
  Erase columns
  ReDim columns(1 To UBound(patchValues, 2))

  Dim c As Long
  For c = LBound(patchValues, 2) To UBound(patchValues, 2)
    If c = idCol Or c = typeCol Then GoTo ContinueCol
    Dim header As String: header = Trim$(CStr(patchValues(1, c)))
    If LenB(header) = 0 Then GoTo ContinueCol

    If StrComp(header, RECOVERY_KEY, vbTextCompare) = 0 Then GoTo ContinueCol
    If Not recovery Is Nothing Then
      If StrComp(header, recovery("InternalName"), vbTextCompare) = 0 Or _
         StrComp(header, recovery("Title"), vbTextCompare) = 0 Or _
         StrComp(header, recovery("EntityPropertyName"), vbTextCompare) = 0 Then GoTo ContinueCol
    End If
    Dim schemaIndex As Long: schemaIndex = FindWritableFieldIndex(header, fields, fieldCount)
    If schemaIndex = 0 Then
      unmatched.Add header
    Else
      colCount = colCount + 1
      columns(colCount).header = header
      columns(colCount).ColIndex = c
      columns(colCount).schemaIndex = schemaIndex
      columns(colCount).InternalName = fields(schemaIndex).InternalName
      columns(colCount).FieldType = fields(schemaIndex).FieldType
      columns(colCount).Required = fields(schemaIndex).Required
      columns(colCount).HasMaxLength = fields(schemaIndex).HasMaxLength
      columns(colCount).MaxLength = fields(schemaIndex).MaxLength
      Set columns(colCount).choices = fields(schemaIndex).choices
    End If
ContinueCol:
  Next

  If colCount = 0 Then
    Erase columns
  ElseIf colCount < UBound(patchValues, 2) Then
    ReDim Preserve columns(1 To colCount)
  End If
End Sub

Private Function FindWritableFieldIndex(ByVal header As String, ByRef fields() As TFieldSchema, ByVal fieldCount As Long) As Long
  Dim target As String: target = LCase$(Trim$(header))
  Dim i As Long
  For i = 1 To fieldCount
    If Not fields(i).Writable Then GoTo ContinueFind
    If LCase$(fields(i).InternalName) = target Or LCase$(fields(i).Title) = target Then
      FindWritableFieldIndex = i
      Exit Function
    End If
ContinueFind:
  Next
End Function

Private Function FieldHeaderName(ByRef field As TFieldSchema) As String
  If LenB(Trim$(field.Title)) > 0 Then
    FieldHeaderName = Trim$(field.Title)
  Else
    FieldHeaderName = Trim$(field.InternalName)
  End If
End Function

Private Function ListHasFieldColumn(ByVal Table As ListObject, ByRef field As TFieldSchema) As Boolean
  Dim col As ListColumn
  For Each col In Table.ListColumns
    Dim header As String: header = Trim$(CStr(col.Name))
    If StrComp(header, field.InternalName, vbTextCompare) = 0 Then
      ListHasFieldColumn = True
      Exit Function
    End If
    If LenB(field.Title) > 0 Then
      If StrComp(header, field.Title, vbTextCompare) = 0 Then
        ListHasFieldColumn = True
        Exit Function
      End If
    End If
  Next
End Function

Private Function FindListColumn(ByVal Table As ListObject, ParamArray names() As Variant) As ListColumn
  Dim col As ListColumn
  For Each col In Table.ListColumns
    Dim n As Long
    For n = LBound(names) To UBound(names)
      If StrComp(Trim$(CStr(col.Name)), CStr(names(n)), vbTextCompare) = 0 Then
        Set FindListColumn = col
        Exit Function
      End If
    Next
  Next
End Function

Private Sub InsertListColumn(ByVal Table As ListObject, ByVal header As String, ByVal position As Long)
  Dim col As ListColumn
  If position < 1 Then position = 1
  If position > Table.ListColumns.Count + 1 Then position = Table.ListColumns.Count + 1

  If Table.ListColumns.Count = 0 Then
    Set col = Table.ListColumns.Add
  Else
    Set col = Table.ListColumns.Add(position)
  End If
  col.Name = header
End Sub

Private Function BuildUnmatchedReason(ByVal unmatched As Collection, ByVal ListTitle As String) As String
  Dim parts As Collection: Set parts = New Collection
  Dim i As Long
  For i = 1 To unmatched.count
    parts.Add "Column '" & CStr(unmatched(i)) & "' does not exist on list '" & ListTitle & "'"
  Next
  BuildUnmatchedReason = JoinCollection(parts, "; ")
End Function

'---------------------------------------
' Row validation + payload build
'---------------------------------------

Private Sub ValidateRow( _
  ByRef patchValues As Variant, _
  ByVal sheetRow As Long, _
  ByVal idCol As Long, _
  ByVal rowOp As ERowOp, _
  ByRef columns() As TColumnMap, _
  ByVal colCount As Long, _
  ByVal updateRowById As Object, _
  ByRef itemId As Long, _
  ByVal issues As Collection _
)
  itemId = 0

  If rowOp = RowOpUpdate Then
    Dim idVal As Variant: idVal = patchValues(sheetRow, idCol)

    If IsError(idVal) Then
      issues.Add "ID is an Excel error"
      Exit Sub
    End If
    If IsEmpty(idVal) Or (VarType(idVal) = vbString And LenB(Trim$(CStr(idVal))) = 0) Then
      issues.Add "ID is blank"
      Exit Sub
    End If
    If Not IsNumeric(idVal) Then
      issues.Add "ID is not numeric"
      Exit Sub
    End If

    itemId = CLng(idVal)
    If itemId <= 0 Then
      issues.Add "ID must be a positive integer"
      Exit Sub
    End If

    Dim idKey As String: idKey = CStr(itemId)
    If updateRowById.Exists(idKey) Then
      issues.Add "Duplicate ID " & idKey
      Exit Sub
    End If
  End If

  Dim i As Long
  For i = 1 To colCount
    Dim raw As Variant: raw = patchValues(sheetRow, columns(i).ColIndex)

    If IsError(raw) Then
      issues.Add "Column '" & columns(i).header & "' contains an Excel error"
      GoTo ContinueValidate
    End If

    If IsBlankCell(raw) Then
      'Blank = skip. On Create, required fields cannot be omitted.
      If rowOp = RowOpCreate And columns(i).Required Then
        issues.Add "Column '" & columns(i).header & "' is required and cannot be blank on Create"
      End If
      GoTo ContinueValidate
    End If

    If IsClearToken(raw) Then
      If columns(i).Required Then
        issues.Add "Column '" & columns(i).header & "' is required and cannot be cleared"
      End If
      GoTo ContinueValidate
    End If

    Dim textVal As String: textVal = Trim$(CStr(raw))
    Call ValidateCellValue(columns(i), raw, textVal, issues)
ContinueValidate:
  Next
End Sub

Private Sub ValidateCellValue( _
  ByRef col As TColumnMap, _
  ByVal raw As Variant, _
  ByVal textVal As String, _
  ByVal issues As Collection _
)
  Select Case CLng(col.FieldType)
    Case SharePointFieldNumber
      If Not IsNumeric(raw) Then
        issues.Add "Column '" & col.header & "' must be numeric"
      End If

    Case SharePointFieldBoolean
      If Not IsBooleanText(textVal) Then
        issues.Add "Column '" & col.header & "' must be a boolean (TRUE/FALSE/1/0/Yes/No)"
      End If

    Case SharePointFieldDate
      If Not IsDate(raw) And Not IsDate(textVal) Then
        issues.Add "Column '" & col.header & "' must be a date"
      End If

    Case SharePointFieldChoice
      If col.choices.count > 0 Then
        If Not col.choices.Exists(textVal) Then
          issues.Add "Column '" & col.header & "' value '" & textVal & "' is not a valid choice"
        End If
      End If

    Case SharePointFieldMultiChoice
      Dim multiChoiceParts() As String: multiChoiceParts = SplitMulti(textVal)
      Dim p As Long
      For p = LBound(multiChoiceParts) To UBound(multiChoiceParts)
        Dim part As String: part = Trim$(multiChoiceParts(p))
        If LenB(part) = 0 Then GoTo ContinueMultiChoice
        If col.choices.count > 0 Then
          If Not col.choices.Exists(part) Then
            issues.Add "Column '" & col.header & "' value '" & part & "' is not a valid choice"
          End If
        End If
ContinueMultiChoice:
      Next

    Case SharePointFieldPerson
      If Not IsPersonValue(textVal) Then
        issues.Add "Column '" & col.header & "' must be a user id or email"
      End If

    Case SharePointFieldMultiPerson
      Dim multiPersonParts() As String: multiPersonParts = SplitMulti(textVal)
      Dim mp As Long
      For mp = LBound(multiPersonParts) To UBound(multiPersonParts)
        Dim personPart As String: personPart = Trim$(multiPersonParts(mp))
        If LenB(personPart) = 0 Then GoTo ContinueMultiPerson
        If Not IsPersonValue(personPart) Then
          issues.Add "Column '" & col.header & "' value '" & personPart & "' must be a user id or email"
        End If
ContinueMultiPerson:
      Next

    Case SharePointFieldLookup
      If Not IsNumeric(raw) Then
        issues.Add "Column '" & col.header & "' must be a lookup id"
      End If

    Case SharePointFieldText
      If col.HasMaxLength And col.MaxLength > 0 Then
        If Len(textVal) > col.MaxLength Then
          issues.Add "Column '" & col.header & "' exceeds MaxLength " & CStr(col.MaxLength)
        End If
      End If

    Case Else
      ' Accept as text-like
  End Select
End Sub

Private Sub CollectRowLogins(ByRef values As Variant, ByVal sheetRow As Long, ByRef columns() As TColumnMap, _
    ByVal colCount As Long, ByVal people As Object, ByVal logins As Collection)
  Dim i As Long, raw As Variant, part As Variant, parts As Variant, login As String
  For i = 1 To colCount
    Select Case columns(i).FieldType
      Case SharePointFieldPerson, SharePointFieldMultiPerson
        raw = values(sheetRow, columns(i).ColIndex)
        If Not IsBlankCell(raw) And Not IsClearToken(raw) Then
          If columns(i).FieldType = SharePointFieldPerson Then
            parts = Array(CStr(raw))
          Else
            parts = SplitMulti(CStr(raw))
          End If
          For Each part In parts
            login = Trim$(CStr(part))
            If LenB(login) > 0 And Not IsNumeric(login) Then
              If Not people.Exists(login) Then
                people.Add login, Empty
                logins.Add login
              End If
            End If
          Next
        End If
    End Select
  Next
End Sub

Private Function BuildRowPayload( _
  ByRef patchValues As Variant, _
  ByVal sheetRow As Long, _
  ByRef columns() As TColumnMap, _
  ByVal colCount As Long, _
  ByVal people As Object, ByVal issues As Collection _
) As stdJSON
  Dim data As stdJSON: Set data = stdJSON.Create(eJSONObject)
  Dim i As Long
  For i = 1 To colCount
    Dim raw As Variant: raw = patchValues(sheetRow, columns(i).ColIndex)
    If IsBlankCell(raw) Then GoTo ContinueBuild ' skip / omit
    If IsClearToken(raw) Then
      data.Add columns(i).InternalName, Null
      GoTo ContinueBuild
    End If

    Dim textVal As String: textVal = Trim$(CStr(raw))
    Select Case CLng(columns(i).FieldType)
      Case SharePointFieldMultiChoice, SharePointFieldMultiPerson
        Dim convertedObj As Variant: Set convertedObj = ConvertCellValue(columns(i), raw, textVal, people, issues)
        data.Add columns(i).InternalName, convertedObj
      Case Else
        Dim converted As Variant: converted = ConvertCellValue(columns(i), raw, textVal, people, issues)
        data.Add columns(i).InternalName, converted
    End Select
    If issues.count > 0 Then Exit Function
ContinueBuild:
  Next
  Set BuildRowPayload = data
End Function

Private Function ConvertCellValue( _
  ByRef col As TColumnMap, _
  ByVal raw As Variant, _
  ByVal textVal As String, _
  ByVal people As Object, ByVal issues As Collection _
) As Variant
  Select Case CLng(col.FieldType)
    Case SharePointFieldNumber, SharePointFieldLookup
      ConvertCellValue = CDbl(raw)
    Case SharePointFieldBoolean
      ConvertCellValue = ParseBoolean(textVal)
    Case SharePointFieldDate
      If IsDate(raw) Then
        ConvertCellValue = CDate(raw)
      Else
        ConvertCellValue = CDate(textVal)
      End If
    Case SharePointFieldMultiChoice, SharePointFieldMultiPerson
      Dim arr As stdJSON: Set arr = stdJSON.Create(eJSONArray)
      Dim parts() As String: parts = SplitMulti(textVal)
      Dim p As Long
      For p = LBound(parts) To UBound(parts)
        Dim part As String: part = Trim$(parts(p))
        If LenB(part) > 0 Then
          If CLng(col.FieldType) = SharePointFieldMultiPerson Then
            arr.Add ResolveUploadPerson(part, col.header, people, issues)
            If issues.count > 0 Then Exit For
          Else
            arr.Add part
          End If
        End If
      Next
      Set ConvertCellValue = arr
    Case SharePointFieldPerson
      ConvertCellValue = ResolveUploadPerson(textVal, col.header, people, issues)
    Case Else
      ConvertCellValue = textVal
  End Select
End Function

Private Function ResolveUploadPerson(ByVal person As String, ByVal header As String, ByVal people As Object, _
    ByVal issues As Collection) As Long
  If IsNumeric(person) Then
    ResolveUploadPerson = CLng(person)
  ElseIf VarType(people(person)) = vbLong Then
    ResolveUploadPerson = people(person)
  Else
    issues.Add header & " (" & person & "): " & people(person)
  End If
End Function

'---------------------------------------
' Results helpers
'---------------------------------------

Private Function BuildEmptyResults(ByRef patchValues As Variant, ByVal idCol As Long, ByVal dataRows As Long) As Variant
  Dim out As Variant: ReDim out(1 To dataRows + 1, 1 To 3)
  out(1, 1) = "ID"
  out(1, 2) = "SuccessFailure"
  out(1, 3) = "FailureReason"

  Dim i As Long
  For i = 1 To dataRows
    Dim idVal As Variant: idVal = patchValues(i + 1, idCol)
    If IsError(idVal) Then
      out(i + 1, 1) = "#ERROR"
    ElseIf IsEmpty(idVal) Then
      out(i + 1, 1) = vbNullString
    Else
      out(i + 1, 1) = idVal
    End If
    out(i + 1, 2) = vbNullString
    out(i + 1, 3) = vbNullString
  Next

  BuildEmptyResults = out
End Function

Private Function ExtractBatchError(ByVal one As stdJSON, ByVal httpStatus As Long) As String
  If one.Exists("data") And IsObject(one("data")) And TypeOf one("data") Is stdJSON Then
    Dim dataObj As stdJSON: Set dataObj = one("data")
    If dataObj.Exists("error") And IsObject(dataObj("error")) And TypeOf dataObj("error") Is stdJSON Then
      Dim errObj As stdJSON: Set errObj = dataObj("error")
      If errObj.Exists("message") Then
        If IsObject(errObj("message")) And TypeOf errObj("message") Is stdJSON Then
          Dim msgObj As stdJSON: Set msgObj = errObj("message")
          If msgObj.Exists("value") Then
            ExtractBatchError = CStr(msgObj("value"))
            Exit Function
          End If
        ElseIf Not IsObject(errObj("message")) Then
          ExtractBatchError = CStr(errObj("message"))
          Exit Function
        End If
      End If
    End If
  End If

  ExtractBatchError = "HTTP " & CStr(httpStatus)
End Function

'---------------------------------------
' Small helpers
'---------------------------------------

Private Function IsClearToken(ByVal raw As Variant) As Boolean
  If IsError(raw) Then Exit Function
  If IsEmpty(raw) Then Exit Function
  If IsNull(raw) Then Exit Function
  If VarType(raw) = vbString Then
    IsClearToken = (StrComp(Trim$(CStr(raw)), CLEAR_TOKEN, vbTextCompare) = 0)
  End If
End Function

Private Function IsBlankCell(ByVal raw As Variant) As Boolean
  If IsEmpty(raw) Then
    IsBlankCell = True
  ElseIf IsNull(raw) Then
    IsBlankCell = True
  ElseIf VarType(raw) = vbString Then
    IsBlankCell = (LenB(Trim$(CStr(raw))) = 0)
  End If
End Function

Private Function IsBooleanText(ByVal textVal As String) As Boolean
  Select Case LCase$(Trim$(textVal))
    Case "true", "false", "1", "0", "yes", "no"
      IsBooleanText = True
  End Select
End Function

Private Function ParseBoolean(ByVal textVal As String) As Boolean
  Select Case LCase$(Trim$(textVal))
    Case "true", "1", "yes"
      ParseBoolean = True
    Case Else
      ParseBoolean = False
  End Select
End Function

Private Function IsPersonValue(ByVal textVal As String) As Boolean
  If IsNumeric(textVal) Then
    IsPersonValue = True
  ElseIf InStr(1, textVal, "@", vbTextCompare) > 0 Then
    IsPersonValue = True
  End If
End Function

Private Function SplitMulti(ByVal textVal As String) As String()
  Dim normalized As String: normalized = Replace(textVal, ",", ";")
  SplitMulti = Split(normalized, ";")
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal delim As String) As String
  Dim sb As String: sb = vbNullString
  Dim i As Long
  For i = 1 To items.count
    If i > 1 Then sb = sb & delim
    sb = sb & CStr(items(i))
  Next
  JoinCollection = sb
End Function

Private Function NzString(ByVal parent As stdJSON, ByVal key As String) As String
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then Exit Function
  If IsNull(parent(key)) Then Exit Function
  NzString = CStr(parent(key))
End Function

Private Function NzBool(ByVal parent As stdJSON, ByVal key As String) As Boolean
  If parent Is Nothing Then Exit Function
  If Not parent.Exists(key) Then Exit Function
  If IsObject(parent(key)) Then Exit Function
  If IsNull(parent(key)) Then Exit Function
  On Error Resume Next
  NzBool = CBool(parent(key))
  On Error GoTo 0
End Function
