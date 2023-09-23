'Dependencies:
'  * stdICallable

'Update a field of a table in Excel to a value.
'@param {ListObject} - The table to filter
'@param {String} - The name of the column to update with information.
'@param {Variant} - The value to update the column to
Public Sub updateFieldConst(ByVal lo As ListObject, ByVal sColumnName As String, ByVal updateTo As Variant)
  lo.ListColumns(sColumnName).DataBodyRange.value = updateTo
End Sub

'Update a field of a table by callback in Excel.
'@param {ListObject} - The table to filter
'@param {String} - The name of the column to update with information.
'@param {stdICallable<(RowObject)=>Variant>} - Callback which will be applied to each row to identify the final value. 1st param will be a RowObject.
'@param {stdICallable<(RowObject)=>Boolean>} - Callback which will be applied to each row to identify whether it should be updated. 1st param will be a RowObject.
'@declare `RowObject` as `Dictionary<string, variant>` - A dictionary representing a row. Dictionary will have field names as keys, paired with value for that row.
'@example `Call updateField(myTable, "fieldName", stdLambda.Create("$1.Category = ""A"""))`
'@requires {stdICallable}
Public Sub updateField(ByVal lo As ListObject, ByVal sColumnName As String, ByVal updateTo As stdICallable, Optional ByVal whereCondition As stdICallable = Nothing, Optional Relationships As Collection = Nothing)
  On Error GoTo EH
  
  Dim iColIndex As Long: iColIndex = lo.ListColumns(sColumnName).index
  Dim iRowCount As Long: iRowCount = lo.ListRows.Count
  Dim vUpdates(): ReDim vUpdates(1 To iRowCount, 1 To 1)
  Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
  Dim rels As Object: If Not Relationships Is Nothing Then Set rels = CreateObject("Scripting.Dictionary")
  Set row("rels__") = rels
  Dim v: v = lo.Range.value
  
  'Scan rows for results...
  Dim i As Long, j As Long
  For i = 2 To UBound(v, 1)
    'Prepare row for callback
    For j = 1 To UBound(v, 2)
      row(v(1, j)) = v(i, j)
    Next
    If Not Relationships Is Nothing Then
        Dim rel As Object
        For Each rel In Relationships
          Set rels(rel("name")) = rel("lookup")(row(rel("fromField")))
        Next
    End If
    
    If whereCondition Is Nothing Then
      vUpdates(i - 1, 1) = updateTo.Run(row)
    Else
      'Update only if where condition succeeds
      If whereCondition.Run(row) Then
        vUpdates(i - 1, 1) = updateTo.Run(row)
      Else
        vUpdates(i - 1, 1) = v(i, iColIndex)
      End If
    End If
  Next
  
  'Update data
  If Not Selection.ListObject Is Nothing Then
    lo.HeaderRowRange.Resize(1, 1).offset(0, lo.ListColumns.Count).Select
  End If
  lo.ListColumns(sColumnName).DataBodyRange.value = vUpdates
  Exit Sub
EH:
  If Err.Description Like "Property or Method *" Then
    Dim sField As String: sField = Split(Mid(Err.Description, 20), " ")(0)
    Err.Raise 1, , "Field '" & sField & "' not present in table"
  End If
End Sub

'Create a relationship object. Links some table's to the `ToTable` where some fields match.
'@param {string}     sRelationshipName - The name of the relationship, this will be the name of the property under rels__.
'@param {string}     sFromField        - The name of the field in the base table which links to the ToTable
'@param {ListObject} loIntoTable       - The table to link to.
'@param {string}     sToField          - The name of the field within the linked table which matches sFromField values
'@param {boolean}    isOneToMany       - If true, returned value will be a stdArray of matches. Default false. 
'@param {Collection} Relationships     - Any sub-relationships for this relationship.
'@example ```
'  Dim mainTableRels as Collection: set mainTableRels = new Collection
'  mainTableRels.add CreateRelationship("Category", Lookups.ListObjects("Categories"), "ID")
'  ...
'  Call updateField(myTable, "CategoryName", stdLambda.Create("$1.rels__.Category.name"), Relationships:= maintTableRels)
'  Call updateField(myTable, "CategoryName", stdLambda.Create("$1.rels__.SubCategory.rels__.Category.name"), Relationships:= maintTableRels)
'```
'@requires {stdICallable}
'@requires {stdArray} Depends on stdArray if isOneToMany
Public Function CreateRelationship(ByVal sRelationshipName as string, ByVal sFromField as string, ByVal loToTable as ListObject, ByVal sToField as string, Optional ByVal isOneToMany as boolean = false, Optional ByVal Relationships as Collection = nothing) as Object
  Dim oRet as object: set oRet = CreateObject("Scripting.Dictionary")
  oRet("name") = sRelationshipName
  oRet("fromField") = sFromField
  oRet("isOneToMany") = isOneToMany
  set oRet("lookup") = CreateObject("Scripting.Dictionary")
  Dim rels As Object: If Not Relationships Is Nothing Then Set rels = CreateObject("Scripting.Dictionary")
  Set oRet("lookup")("rels__") = rels

  'Create relationship
  Dim v: v = loToTable.Range.Value
  Dim i as long, j as long
  For i = 2 to ubound(v,1)
    'Create and populate row object
    Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(v, 2)
      row(v(1, j)) = v(i, j)
    Next
    If Not Relationships Is Nothing Then
        Dim rel As Object
        For Each rel In Relationships
          Set rels(rel("name")) = rel("lookup")(row(rel("fromField")))
        Next
    End If

    'Produce lookups
    Dim sID as string: sID = row(sToField)
    if isOneToMany then
      if not oRet("lookup").exists(sID) then set oRet("lookup")(sID) = stdArray.Create()
      Call oRet("lookup").item(sID).push(row) 
    else
      set oRet("lookup")(sID) = row
    end if
  next
  set CreateRelationship = oRet
End Function

'Applies a filters to a table by callback in Excel. The table requires an ID field.
'@param {ListObject} - The table to filter
'@param {String} - The name of the ID field of the table. This will be used to apply the filter.
'@param {stdICallable<(RowObject)=>Boolean>} - Callback which will be applied to each row. 1st param will be a RowObject.
'@declare `RowObject` as `Dictionary<string, variant>` - A dictionary representing a row. Dictionary will have field names as keys, paired with value for that row.
'@example `Call filterTable(myTable, "IDField", stdLambda.Create("$1.Category = ""A"""))`
'@dependency {stdICallable}
Public Sub applyFilterToTable(ByVal lo As ListObject, ByVal idField As String, ByVal callback As stdICallable)
  Dim sIDs() As String: ReDim sIDs(1 To lo.ListRows.Count)
  Dim iIndex As Long: iIndex = 1
  Dim iIDField As Long: iIDField = lo.ListColumns(idField).index
  Dim v: v = lo.Range.value
  Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
  Dim i As Long, j As Long
  
  'Scan rows for results...
  For i = 1 To UBound(v, 1)
    'Prepare row for callback
    For j = 1 To UBound(v, 2)
      row(v(1, j)) = v(i, j)
    Next
    
    'Call callback on row, if succeeds assign ID to result
    If callback.Run(row) Then
      sIDs(iIndex) = v(i, iIDField)
      iIndex = iIndex + 1
    End If
  Next
  
  'Apply filter
  If sIDs(1) <> Empty Then
    If UBound(sIDs) > 1 Then ReDim Preserve sIDs(1 To iIndex - 1)
    lo.AutoFilter.ShowAllData
    lo.Range.AutoFilter iIDField, sIDs, xlFilterValues
  else
    lo.Range.AutoFilter iIDField, Array("9e49913b-3172-4a7e-b377-07a9ff4369ed"), xlFilterValues
  End If
End Sub