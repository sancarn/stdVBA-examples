VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@dependencies
'  * stdEnumerator
'  * stdCallback
'  * stdICallable
'  * stdJSON - For JSON parsing and exporting
'  * stdCSV - For CSV parsing and exporting
'@devnote - You will likely need `stdLambda` to interface with the library effectively.

'
'Spec:
'Constructors:
'    CreateFromRecordSet()
'    CreateFromListObject()
'    CreateFromArray()
'    CreateFromEnumerator()
'Instance Methods:
'    select(whereLambda) as stdTable
'    update(updateLambda) as stdTable
'    updateField(sFieldName, vValue) as stdTable
'    join(table2,field1,field2,joinType) as stdTable
'
' Dim myTable as stdTable
' set myTable = stdTable.CreateFromListObject(Sheets(1).ListObjects(1))
' myTable.select(stdLambda.Create("$1.item(""FieldName"")=""Value""")) _
'        .updateField("SomeField","SomeValue")
'        .update(stdLambda.CreateMultiline(Array( _
'          "let $1.item(""SomeField1"")=""Value1""", _
'          "let $1.item(""SomeField2"")=""Value2""", _
'        )))
'        .join("alias", otherTable,"FieldOrLambda", "FieldOrLambda",stdTableJoinType.Inner)
'        .filter(stdLambda.Create("$1!join!alias!"Some Field" = ""Value""")
'        .filter(stdLambda.Create("[Some field] = ""value"""))
'        .addField(...)
'        .removeField(...)
'        .fields  'fields ==> stdEnumerator<Dictionary>
'        .addCalculatedField(sFileName, calculationLambda)
'        .rows.map(mapper) 'rows ==> stdEnumerator<stdTableRow>
'Integration examples:
'*  Call stdJSON.CreateFromVBObject(This.Rows.ToCollection()).ToFile(path)

Private Type TThis
  name As String
  headers() As String
  rows As stdEnumerator
End Type
Private This As TThis

'Create a new table from a 2d array
'@param name - The name of the table
'@param rows - The array to create the table from
'@param headers as Variant<Array<String>> - The headers of the table. If not specified, the first row of the array is used.
'@returns - A new table containing the data from the array
Public Function Create(ByVal name As String, ByVal rows As stdEnumerator, ByVal headers As Variant) As stdTable
  Set Create = New stdTable
  ReDim Preserve headers(1 To UBound(headers) - LBound(headers) + 1) 'Make headers array 1-based
  Call Create.protInit(name, headers, rows)
End Function

'Create a new table from a collection
'@param c as Collection - The collection to create the table from
'@param Optional tableName as string - The name of the table. If not specified, "Collection" is used.
'@param Optional fieldName as string - The name of the field to use. If not specified, "Value" is used.
'@returns - A new table containing the data from the collection
Public Function CreateFromCollection(ByVal c As Collection, Optional ByVal tableName As String = "Collection", Optional ByVal fieldName As String = "Value", Optional autoExpand As Boolean = True) As stdTable
  Set CreateFromCollection = CreateFromVariantList(c, tableName, fieldName, autoExpand)
End Function

'Create a new table from a array
'@param c as Variand<Array1D<Variant>> - The 1D array to create the table from
'@param Optional tableName as string - The name of the table. If not specified, "Collection" is used.
'@param Optional fieldName as string - The name of the field to use. If not specified, "Value" is used.
'@returns - A new table containing the data from the collection
Public Function CreateFromArray1D(ByVal c As Collection, Optional ByVal tableName As String = "Collection", Optional ByVal fieldName As String = "Value", Optional autoExpand As Boolean = True) As stdTable
  Set CreateFromArray1D = CreateFromVariantList(c, tableName, fieldName, autoExpand)
End Function

'Create from variant list
Private Function CreateFromVariantList(ByVal c As Variant, Optional ByVal tableName As String = "Collection", Optional ByVal fieldName As String = "Value", Optional autoExpand As Boolean = True) As stdTable
  'Map collection to collection of rows
  Dim cc As Collection: Set cc = New Collection
  Dim v: For Each v In c
    Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
    row.add fieldName, v
    Call cc.add(row)
  Next

  'Create table
  Set CreateFromVariantList = Create(tableName, stdEnumerator.CreateFromIEnumVariant(cc), Array(fieldName))
  If autoExpand Then
    Call CreateFromVariantList.FieldExpand(fieldName)
    Call CreateFromVariantList.FieldsRemove(fieldName)
  End If
End Function

'Create a new table from a list object
'@param lo as ListObject - The list object to create the table from
'@returns - A new table containing the data from the list object
Public Function CreateFromListObject(ByVal lo As ListObject) As stdTable
  Dim headers() As String: ReDim headers(1 To lo.ListColumns.Count)
  Dim vHeaders: vHeaders = lo.HeaderRowRange.value
  With lo
    For i = 1 To .ListColumns.Count
      headers(i) = vHeaders(1, i)
    Next i
  End With
  
  Set CreateFromListObject = New stdTable
  Call CreateFromListObject.protInit(lo.name, headers, stdEnumerator.CreateFromListObject(lo))
End Function

'Create a new table from a table name
'@param tableName - The name of the table to create the table from
'@param ws - The worksheet to search for the table in. If not specified, all worksheets in the workbook are searched.
'@param wb - The workbook to search for the table in. If not specified, the current workbook is searched.
'@returns - A new table containing the data from the list object
Public Function CreateFromTableByName(ByVal tableName As String, Optional ByVal ws As Worksheet = Nothing, Optional ByVal wb As Workbook = Nothing) As stdTable
  If wb Is Nothing Then Set wb = ThisWorkbook
  Dim lo As ListObject
  If ws Is Nothing Then
    For Each ws In wb.worksheets
      On Error Resume Next
      Set lo = ws.ListObjects(tableName)
      If Not lo Is Nothing Then Exit For
    Next
  Else
    Set lo = ws.ListObjects(tableName)
  End If
  If lo Is Nothing Then Err.Raise 5, "stdTable.CreateFromTableByName", "Table '" & tableName & "' not found."
  Set CreateFromTableByName = CreateFromListObject(lo)
End Function

'Create a new table from a recordset
'@param rs as Object<ADODB.RecordSet> - The recordset to create the table from
'@returns stdTable - A new table containing the data from the recordset
Public Function CreateFromRecordSet(ByVal rs As Object) As stdTable
  Dim headers() As String: ReDim headers(1 To rs.fields.Count)
  Dim i As Long
  For i = 1 To rs.fields.Count
    headers(i) = rs.fields(i).name
  Next i

  Dim getNext As stdCallback: Set getNext = stdCallback.CreateFromObjectMethod(Me, "protGetDictFromRecordSet").Bind(rs)

  Set CreateFromRecordSet = New stdTable
  Call CreateFromRecordSet.protInit(rs.name, headers, stdEnumerator.CreateFromCallable(getNext))
End Function

'Create a table from a 2d array
'@param name - The name of the table
'@param vArr as Variant<Array2D<Variant>> - The array to create the table from
'@param Optional headers as Boolean - If true, the first row of the array is used as the headers. If false, the headers are not set. If an array of strings, the headers are set to numbers.
'@returns - A new table containing the data from the array
Public Function CreateFromArray2D(ByVal name As String, ByVal vArr As Variant, Optional ByVal headers As Boolean = True) As stdTable
  Dim c As Collection: Set c = New Collection
  Dim vHeaders(): ReDim vHeaders(1 To UBound(vArr, 2) - LBound(vArr, 2) + 1)
  Dim i As Long
  For i = 1 To UBound(vArr, 2)
    If headers Then
      vHeaders(i) = vArr(1, i)
    Else
      vHeaders(i) = CStr(i)
    End If
  Next i
  
  Dim j As Long
  For i = 1 To UBound(vArr, 1)
    Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(vArr, 2)
      If headers Then
        row.add vArr(1, j), vArr(i, j)
      Else
        row.add CStr(j), vArr(i, j)
      End If
    Next
    Call c.add(row)
  Next

  Set CreateFromArray2D = Create(name, stdEnumerator.CreateFromIEnumVariant(c), vHeaders)
End Function

'Create a table from an excel range
'@param rng as Object<Excel.Range> - The range to create the table from
'@param Optional headers as Boolean - If true, the first row of the range is used as the headers. If false, the headers are not set. If an array of strings, the headers are set to numbers.
'@returns - A new table containing the data from the range
Public Function CreateFromExcelRange(ByVal rng As Object, Optional ByVal headers As Boolean = True) As stdTable
  Set CreateFromExcelRange = CreateFromArray2D(rng.address, rng.value, headers)
End Function

'Create a new dictionary from a recordset
'@protected
'@param rs as Object<ADODB.RecordSet> - The recordset to create the dictionary from
'@param lastItem as Variant<Object<ADODB.RecordSet>|null> - The last dict returned.
'@param index as long - The index of the item to be returned from the recordset. If this is the first call, pass 1.
'@returns Dictionary<string,variant> - A new dictionary containing the data from the recordset
Public Function protGetDictFromRecordSet(ByVal rs As Object, ByVal lastItem As Variant, ByVal index As Long) As Variant
  If index > 1 Then Call rs.MoveNext
  If rs.EOF Then
    protGetDictFromRecordSet = Null
    Exit Function
  End If

  Set protGetDictFromRecordSet = CreateObject("Scripting.Dictionary")
  protGetDictFromRecordSet.add "__index", index
  Dim i As Long
  For i = 0 To rs.fields.Count - 1
    protGetDictFromRecordSet.add rs.fields(i).name, rs.fields(i).value
  Next i
End Function

'Get a field of an object
'@protected
'@param field - The field to get
'@param dict as Object<Dictionary<string,variant>> - The object to get the field from
'@returns Variant - The value of the field
Public Function protGetDictionaryField(ByVal field As String, ByVal dict As Object, ByVal keyToIgnore As Variant) As Variant
  Call CopyVariant(protGetDictionaryField, dict.item(field))
End Function

'Initialize the current table
'@param name as string - The name of the table
'@param headers as Variant - The headers of the table
'@param rows as stdEnumerator - The rows of the table
Public Sub protInit(name As String, headers As Variant, rows As stdEnumerator)
  With This
    .name = name
    .headers = headers
    Set .rows = rows
  End With
End Sub


'Get/Set the name of the table
Public Property Get name() As String
  name = This.name
End Property
Public Property Let name(ByVal value As String)
  This.name = value
End Property

'Get the headers of the table
Public Property Get headers() As Variant
  headers = This.headers
End Property

'Get the rows of the table
'@returns stdEnumerator<Dictionary<string,variant>> - An enumerator containing the rows of the table
Public Property Get rows() As stdEnumerator
  Set rows = This.rows
End Property

'Apply a lambda to each row of the current table
'@param lambda as stdICallable<(row:Object)=>Void> - A lambda or callback that is applied to each row. This can be used to update the rows in place.
'@returns stdTable - The existing table, including any mutated rows
Public Function ForEach(ByVal lambda As stdICallable) As stdTable
  Set update = Me
  Call This.rows.ForEach(lambda)
End Function

'Filter the current table using the specified lambda
'@param lambda as stdICallable<(row:Object)=>boolean> - A lambda or callback that takes a row and returns a boolean. If true, the row is kept.
'@returns stdTable - A new table containing only the rows that passed the filter
Public Function Filter(ByVal lambda As stdICallable) As stdTable
  Set Filter = New stdTable
  Call Filter.protInit(This.headers, This.rows.Filter(lambda))
End Function

'Creates a new table containing only the specified fields from the current table
'@param Fields() as string - The fields to select
'@returns stdTable - A new table containing only the specified fields
Public Function FieldsSelect(ParamArray fields()) As stdTable
  Set FieldsSelect = Me
  ReDim Preserve fields(1 To UBound(fields) + 1)
  This.headers = fields
End Function

'Remove fields from the current table
'@param ParamArray<> - The fields to remove
'@returns stdTable - The existing table with the fields removed
Public Function FieldsRemove(ParamArray fields()) As stdTable
  Set FieldsRemove = Me
  Dim toRemove As Object: Set toRemove = CreateObject("Scripting.Dictionary")
  Dim i As Long
  For i = 0 To UBound(fields)
    toRemove.add fields(i), True
  Next i

  'Remove from headers
  Dim vNewHeaders() As String: ReDim vNewHeaders(1 To UBound(This.headers) - UBound(fields))
  Dim index As Long: index = 1
  For i = 1 To UBound(This.headers)
    If Not toRemove(This.headers(i)) Then
      vNewHeaders(index) = This.headers(i)
      index = index + 1
    End If
  Next
  This.headers = vNewHeaders
End Function

'Update a field of the current table using the specified lambda
'@param field - The name of the field to update
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to set the field to
'@returns stdTable - The existing table.
Public Function FieldUpdate(ByVal field As String, ByVal lambda As stdICallable) As stdTable
  Set FieldUpdate = Me
  Dim row As Object
  For Each row In This.rows
    row(field) = lambda.Run(row)
  Next
End Function

'Rename fields
'@param Paramarray<> - The fields to rename. Each pair of arguments is a field to rename and the new name for that field.
'@returns stdTable - The existing table with mutated field names
Public Function FieldsRename(ParamArray fields()) As stdTable
  Set FieldsRename = Me

  Dim i As Long
  For i = 0 To UBound(fields) Step 2
    Dim oldName As String: oldName = fields(i)
    Dim newName As String: newName = fields(i + 1)
    Dim j As Long, renamed As Boolean: renamed = False
    For j = 1 To UBound(This.headers)
      If This.headers(j) = oldName Then
        This.headers(j) = newName
        renamed = True
        Exit For
      End If
    Next j
    If Not renamed Then Err.Raise 5, "stdTable.FieldsRename", "Field '" & oldName & "' not found"
  Next i

  Dim row As Object
  For Each row In This.rows
    For i = 0 To UBound(fields) Step 2
      row.item(fields(i + 1)) = row.item(fields(i))
      Call row.remove(fields(i))
    Next i
  Next row
End Function

'Add a field to the current table
'@param field as string - The name of the field to add
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to set the field to
'@returns stdTable - The existing table with the new field
Public Function FieldAdd(ByVal field As String, ByVal lambda As stdICallable) As stdTable
  Set FieldAdd = Me
  
  'Ensure field doesn't already exist
  Dim i As Long: For i = 1 To UBound(This.headers)
    If This.headers(i) = field Then Err.Raise 5, "stdTable.FieldAdd", "Field '" & field & "' already exists"
  Next

  'Add to fields
  ReDim Preserve This.headers(1 To UBound(This.headers) + 1)
  This.headers(UBound(This.headers)) = field

  'Map lambda to field
  Dim row As Object
  For Each row In This.rows
    row(field) = lambda.Run(row)
  Next
End Function

'Expand object fields
'@param field as string - The field to expand
'@returns stdTable - The existing table.
'@example - For the table `[{"Car":{"Doors":2}}, {"Car":{"Doors":3}}, ...]`
'           The code `FieldExpand("Car")` would return the table `[{"Doors":2,"Car":{"Doors":2}}, {"Doors":3,"Car":{"Doors":2}}, ...]`
Public Function FieldExpand(ByVal field As String) As stdTable
  Set FieldExpand = Me
  For Each row In This.rows
    If IsObject(row(field)) Then
      Dim target As Object: Set target = row(field)
      Select Case TypeName(target)
        Case "Dictionary"
          Dim key As Variant
          For Each key In target.keys()
            Call CopyVariant(row(key), target(key))
          Next
        Case Else
          Dim prop As Variant
          For Each prop In stdCOM.Create(target).Properties
            Call CopyVariant(row(prop), stdCallback.CreateFromObjectProperty(target, prop, VbGet).Run())
          Next
      End Select
    Else
      Err.Raise 5, "stdTable.FieldExpand", "Field '" & field & "' does not contain an object. Cannot expand."
    End If
  Next
End Function

'Update a field of the current table to a specified value
'@param field - The name of the field to update
'@returns stdTable - The existing table.
Public Function FieldUpdateStatic(ByVal field As String, ByVal value As Variant) As stdTable
  Set FieldUpdateStatic = Me
  Dim row As Object
  For Each row In This.rows
    row(field) = value
  Next
End Function

'Link 1 table to another
'@param table2 as stdTable - The table to join to
'@param field1 - The field in the current table to join on
'@param field2 - The field in the other table to join on
'@param joinAlias - The alias to in the join field
'@param returnOne - If true, only return the first matching row from the other table. Otherwise, return all matching rows. If no rows match, return Null if returnOne is true, else return an empty stdEnumerator.
'@returns stdTable - The existing table with mutated join values
Public Function Join(ByVal table2 As stdTable, ByVal field1 As String, ByVal field2 As String, Optional ByVal joinAlias As String = "", Optional ByVal returnOne As Boolean = False) As stdTable
  Set Join = Me
  Dim groups As Object: Set groups = table2.rows.GroupBy(stdCallback.CreateFromObjectMethod(Me, "protGetDictionaryField").Bind(field2))
  If returnOne Then
    Dim vKey: For Each vKey In groups.keys()
      groups(vKey) = groups(vKey).item(1)
    Next
  End If
  For Each obj In This.rows
    If obj("join") Is Nothing Then Set obj("join") = CreateObject("Scripting.Dictionary")
    If joinAlias = "" Then joinAlias = table2.name
    Dim sKey As String: sKey = obj(field1)
    If groups.Exists(sKey) Then
      Set obj("join")(joinAlias) = groups(sKey)
    Else
      If returnOne Then
        obj("join")(joinAlias) = Null
      Else
        obj("join")(joinAlias) = stdEnumerator.CreateEmpty()
      End If
    End If
  Next
End Function

'Concat rows from a Collection to the current table
'@param c as Collection - The collection to concat
'@returns stdTable - The existing table with the new rows
'@remark - The collection must contain objects with the same fields as the current table. AddRow() can be used to add a single row but is significantly slower than using a collection.
Public Function Concat(ByVal c As Collection) As stdTable
  Set Concat = New stdTable
  Call Concat.protInit(This.name, This.headers, This.rows.Concat(c))
End Function

'Add a row to the current table
'@param ParamArray<Variant> - The values to add to the row. The number of values must match the number of fields in the table.
'@returns stdTable - The existing table with the new row
'@remark - Use Concat() to add multiple rows. This method will be significantly slower.
Public Function AddRow(ParamArray values()) As stdTable
  Set AddRow = Me
  If UBound(values) <> UBound(This.headers) Then Err.Raise 5, "stdTable.AddRow", "Number of parameters must match number of fields"
  Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
  Dim i As Long: For i = 1 To UBound(This.headers)
    row.add This.headers(i), values(i)
  Next
  Dim c As New Collection: Call c.add(row)
  Set AddRow = Concat(c)
End Function

'Reverse the order of the rows in the current table
'@returns stdTable - The existing table with the rows reversed
Public Function Reverse() As stdTable
  Set Reverse = New stdTable
  Call Reverse.protInit(This.name, This.headers, This.rows.Reverse())
End Function

'Group the current table by the specified field. Each row generated by this function will have 2 fields, 1 real and 1 virtual. The virtual field will not be part of the object but not in the headers (and not export in any To___() methods).
'Output is of the form: `[{key:..., group:stdEnumerator<Dictionary<string,variant>>},...]`
'@param fieldName as string - The name of the field to group by
'@param virtualName as string - The name of the virtual field to store the grouped rows in
'@returns - A new table containing the grouped rows
'@remark - This function uses `GroupBy` internally.
Public Function GroupByField(ByVal fieldName As String, Optional ByVal virtualName As String = "group") As stdTable
  Set GroupByField = GroupBy(stdCallback.CreateFromObjectMethod(Me, "protGetDictionaryField").Bind(fieldName), fieldName, virtualName)
End Function

'Group the current table by the specified lambda. Each row generated by this function will have 2 fields, 1 real and 1 virtual. The virtual field will not be part of the object but not in the headers (and not export in any To___() methods).
'Output is of the form: `[{key:..., group:stdEnumerator<Dictionary<string,variant>>},...]`
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to group by
'@param keyName as string - The name of the field to store the key in
'@param virtualName as string - The name of the virtual field to store the grouped rows in
'@returns - A new table containing the grouped rows
Public Function GroupBy(ByVal lambda As stdICallable, Optional ByVal keyName As String, Optional ByVal virtualName As String = "group") As stdTable
  Set GroupBy = New stdTable
  Set groups = This.rows.GroupBy(lambda)
  Dim cRows As Collection: Set cRows = New Collection
  Dim vKey: For Each vKey In groups.keys()
    Dim newRow As Object: Set newRow = CreateObject("Scripting.Dictionary")
    newRow.add keyName, vKey
    newRow.add virtualName, groups(vKey)
    cRows.add newRow
  Next
  Dim headers() As String: ReDim headers(1 To 1)
  headers(1) = keyName
  Call GroupBy.protInit(This.name, headers, stdEnumerator.CreateFromIEnumVariant(cRows))
End Function

'Deep copy the current table
'@returns stdTable - A new table with the same data as the current table
Public Function Clone() As stdTable
  Set Clone = New stdTable
  Dim cRows As Collection: Set cRows = New Collection
  Dim row As Object
  For Each row In This.rows
    Dim newRow As Object: Set newRow = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In This.headers
      newRow.add key, row(key)
    Next
    newRow.add "join", row("join") 'maintain joins
    cRows.add newRow
  Next
  Call Clone.protInit(This.name, This.headers, stdEnumerator.CreateFromIEnumVariant(cRows))
End Function

'Remove duplicates based on some generated id
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to group by to find unique values
'@returns stdTable - A new table with the duplicates removed
Public Function Unique(ByVal lambda As stdICallable) As stdTable
  Set Unique = New stdTable
  Call Unique.protInit(This.name, This.headers, This.rows.Unique(lambda))
End Function

'Export the table to a 2D variant array
'@returns Variant - A 2D variant array containing the data from the table
Public Function ToArray2D() As Variant
  Dim arr() As Variant: ReDim arr(1 To This.rows.length + 1, 1 To UBound(This.headers))
  Dim j As Long
  For j = 1 To UBound(This.headers)
    arr(1, j) = This.headers(j)
  Next j
  Dim i As Long: i = 1
  For Each row In This.rows
    i = i + 1
    For j = 1 To UBound(This.headers)
      arr(i, j) = row(This.headers(j))
    Next j
  Next row
  ToArray2D = arr
End Function

'Export the table to a list object
'@param rng as Range - The range to export the table to
'@param Optional tableName as string - The name of the table. If not specified, this table's name is used.
'@param Optional insertRows as boolean - If true, insert rows into the worksheet. Otherwise, overwrite the existing range.
'@returns ListObject - The list object that was created
'@remark - If a list object already exists in the specified range, it will be overwritten.
Public Function ToListObject(ByVal rng As Range, Optional ByVal tableName As String = "", Optional ByVal insertRows As Boolean = False, Optional ByVal useCurrentRegion As Boolean = True) As ListObject
  Dim arr() As Variant: arr = ToDataBodyArray()
  Dim iUB1 As Long: iUB1 = UBound(arr, 1) + 1 '+1 for headers
  Dim iUB2 As Long: iUB2 = UBound(arr, 2)
  Dim ws As Worksheet: Set ws = rng.Worksheet
  Dim lo As ListObject: Set lo = rng.ListObject
  If Not lo Is Nothing Then Set rng = lo.Range.resize(1, 1)
  If tableName = "" Then
    If lo Is Nothing Then
      tableName = This.name
    Else
      tableName = lo.name
    End If
  End If

  'Calculate original range
  Dim originalRange As Object
  If Not lo Is Nothing Then
    Set originalRange = lo.Range
  ElseIf useCurrentRegion And (rng.currentRegion.rows.Count > 1 Or rng.currentRegion.columns.Count > 1) Then
    Set originalRange = rng.currentRegion
  Else
    Set originalRange = rng
  End If

  'Insert rows if needed
  If insertRows Then
    If originalRange.rows.Count < iUB1 Then Call originalRange.offset(originalRange.rows.Count).resize(originalRange.rows.Count - iUB1).insert(xlShiftDown)
    If originalRange.columns.Count < iUB2 Then Call originalRange.offset(, originalRange.columns.Count).resize(, originalRange.columns.Count - iUB2).insert(xlShiftRight)
  End If

  'resize original range and existing listobject to match array
  Dim hRng As Range: Set hRng = originalRange.resize(1, iUB2)
  Dim dRng As Range: Set dRng = originalRange.offset(1).resize(iUB1 - 1, iUB2)
  Dim loRng As Range: Set loRng = originalRange.resize(iUB1, iUB2)
  If Not lo Is Nothing Then Call lo.resize(loRng)
  hRng.value2 = This.headers
  dRng.value2 = arr
  
  'create listobject if needed
  If lo Is Nothing Then
    Set ToListObject = loRng.Parent.ListObjects.add(xlSrcRange, loRng, xlListObjectHasHeaders:=xlYes)
    ToListObject.name = tableName
  Else
    Set ToListObject = lo
  End If
End Function

'Export the table to a recordset
'@returns Object<ADODB.RecordSet> - An ADODB recordset containing the data from the table
Public Function toRecordSet() As Object
  Const adLockPessimistic As Long = 2
  Const adVariant As Long = 12
  Dim rs As Object: Set rs = CreateObject("ADODB.RecordSet")
  Dim i As Long
  rs.fields.Append "_ID", adInteger
  For i = 1 To UBound(This.headers)
    rs.fields.Append This.headers(i), adVariant
  Next i
  rs.LockType = adLockPessimistic
  rs.open
  Dim row As Object
  i = 0
  For Each row In This.rows
    rs.AddNew
    i = i + 1
    rs.fields("ID").value = i
    For i = 1 To UBound(This.headers)
      rs.fields(This.headers(i)).value = row(This.headers(i))
      rs.update
    Next i
  Next
  Set toRecordSet = rs
End Function

'Export the table to a collection of dictionaries
'@returns Collection<Dictionary<string,variant>> - A collection of dictionaries containing the data from the table
'@remark - Strips virtual fields
Public Function ToCollection() As Collection
  Dim c As Collection: Set c = New Collection
  Dim row As Object
  For Each row In This.rows
    Dim rowNew As Object: Set rowNew = CreateObject("Scripting.Dictionary")
    Dim vHeader
    For Each vHeader In This.headers
      rowNew.add vHeader, row(vHeader)
    Next
    c.add row
  Next
  Set ToCollection = c
End Function

'Export table to JSON
'@returns - A JSON string containing the data from the table
'@requires stdJSON
Public Function toJSON() As String
  With stdJSON.Create(eJSONArray)
    Dim row As Object
    For Each row In This.rows.ToCollection()
      Call .add(row)
    Next
    toJSON = .ToString()
  End With
End Function







'Copy a variant from src to destination
'@param dest - The destination variant
'@param src - The source variant
Private Sub CopyVariant(ByRef dest As Variant, ByVal src As Variant)
  If IsObject(src) Then
    Set dest = src
  Else
    Let dest = src
  End If
End Sub

'Export the table to a 2D variant array
'@returns Variant - A 2D variant array containing the data from the table
Private Function ToDataBodyArray() As Variant
  Dim arr() As Variant: ReDim arr(1 To This.rows.length, 1 To UBound(This.headers))
  Dim i As Long: i = 0
  For Each row In This.rows
    i = i + 1
    Dim j As Long
    For j = 1 To UBound(This.headers)
      arr(i, j) = row(This.headers(j))
    Next j
  Next row
  ToDataBodyArray = arr
End Function
