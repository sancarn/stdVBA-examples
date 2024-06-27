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
  Name as string
  Headers() as String
  Rows as stdEnumerator
End Type
Private This as TThis

'Create a new table from a 2d array
'@param name - The name of the table
'@param rows - The array to create the table from
'@param headers as Variant<Array<String>> - The headers of the table. If not specified, the first row of the array is used.
'@returns - A new table containing the data from the array
Public Function Create(ByVal name as string, ByVal rows as stdEnumerator, ByVal headers as Variant) as stdTable
  set Create = new stdTable
  Redim Preserve headers(1 to UBound(headers)-Lbound(headers)+1) 'Make headers array 1-based
  Call Create.protInit(name, headers, rows)
End Function

'Create a new table from a collection
'@param c as Collection - The collection to create the table from
'@param Optional tableName as string - The name of the table. If not specified, "Collection" is used.
'@param Optional fieldName as string - The name of the field to use. If not specified, "Value" is used.
'@returns - A new table containing the data from the collection
Public Function CreateFromCollection(ByVal c as Collection, optional byval tableName as string = "Collection", optional byval fieldName as string = "Value", optional autoExpand as boolean = true) as stdTable
  set CreateFromCollection = CreateFromVariantList(c, tableName, fieldName, autoExpand)
End Function

'Create a new table from a array
'@param c as Variand<Array1D<Variant>> - The 1D array to create the table from
'@param Optional tableName as string - The name of the table. If not specified, "Collection" is used.
'@param Optional fieldName as string - The name of the field to use. If not specified, "Value" is used.
'@returns - A new table containing the data from the collection
Public Function CreateFromArray1D(ByVal c as Collection, optional byval tableName as string = "Collection", optional byval fieldName as string = "Value", optional autoExpand as boolean = true) as stdTable
  set CreateFromArray1D = CreateFromVariantList(c, tableName, fieldName, autoExpand)
End Function

'Create from variant list
Private Function CreateFromVariantList(ByVal c as variant, optional byval tableName as string = "Collection", optional byval fieldName as string = "Value", optional autoExpand as boolean = true) as stdTable
  'Map collection to collection of rows
  Dim cc as Collection: set cc = new Collection
  Dim v: For each v in c
    Dim row as object: set row = CreateObject("Scripting.Dictionary")
    row.add fieldName, v
    Call cc.Add(row)
  next

  'Create table
  set CreateFromVariantList = Create(tableName, stdEnumerator.CreateFromIEnumVariant(cc), Array(fieldName))
  if autoExpand then 
    Call CreateFromVariantList.FieldExpand(fieldName)
    Call CreateFromVariantList.FieldsRemove(fieldName)
  end if
End Function

'Create a new table from a list object
'@param lo as ListObject - The list object to create the table from
'@returns - A new table containing the data from the list object
Public Function CreateFromListObject(ByVal lo as ListObject) as stdTable
  Dim headers() as String: Redim headers(1 to lo.ListColumns.Count)
  Dim vHeaders: vHeaders = lo.HeaderRowRange.Value
  With lo
    For i = 1 to .ListColumns.Count
      headers(i) = vHeaders(1,i)
    Next i
  End With
  
  set CreateFromListObject = new stdTable
  Call CreateFromListObject.protInit(lo.name, headers, stdEnumerator.CreateFromListObject(lo))
End Function

'Create a new table from a table name
'@param tableName - The name of the table to create the table from
'@param ws - The worksheet to search for the table in. If not specified, all worksheets in the workbook are searched.
'@param wb - The workbook to search for the table in. If not specified, the current workbook is searched.
'@returns - A new table containing the data from the list object
Public Function CreateFromTableByName(ByVal tableName as string, Optional ByVal ws as worksheet = nothing, optional byval wb as workbook = nothing) as stdTable
  if wb is nothing then set wb = ThisWorkbook
  dim lo as ListObject
  if ws is nothing then
    for each ws in wb.worksheets
      On Error Resume Next
      set lo = ws.ListObjects(tableName)
      if not lo is nothing then exit for
    next
  else
    set lo = ws.ListObjects(tableName)
  end if
  if lo is nothing then Err.Raise 5, "stdTable.CreateFromTableByName", "Table '" & tableName & "' not found."
  set CreateFromTableByName = CreateFromListObject(lo)
End Function

'Create a new table from a recordset
'@param rs as Object<ADODB.RecordSet> - The recordset to create the table from
'@returns stdTable - A new table containing the data from the recordset
Public Function CreateFromRecordSet(ByVal rs as Object) as stdTable
  Dim headers() as String: Redim headers(1 to rs.Fields.Count)
  Dim i as long
  For i = 1 to rs.Fields.Count
    headers(i) = rs.Fields(i).Name
  Next i

  dim getNext as stdCallback: set getNext = stdCallback.CreateFromObjectMethod(me, "protGetDictFromRecordSet").bind(rs)

  set CreateFromRecordSet = new stdTable
  Call CreateFromRecordSet.protInit(rs.name, headers,stdEnumerator.CreateFromCallable(getNext))
End Function

'Create a table from a 2d array
'@param name - The name of the table
'@param vArr as Variant<Array2D<Variant>> - The array to create the table from
'@param Optional headers as Boolean - If true, the first row of the array is used as the headers. If false, the headers are not set. If an array of strings, the headers are set to numbers.
'@returns - A new table containing the data from the array
Public Function CreateFromArray2D(ByVal name as string, ByVal vArr as Variant, Optional ByVal headers as boolean = true) as stdTable
  Dim c as Collection: set c = new Collection
  Dim vHeaders(): Redim vHeaders(1 to UBound(vArr,2)-Lbound(vArr,2)+1)
  Dim i as long
  For i = 1 to UBound(vArr,2)
    if headers then 
      vHeaders(i) = vArr(1,i)
    else
      vHeaders(i) = CStr(i)
    end if
  Next i
  
  Dim j as long
  For i = 1 to ubound(vArr,1)
    Dim row as object: set row = CreateObject("Scripting.Dictionary")
    For j = 1 to ubound(vArr,2)
      if headers then 
        row.Add vArr(1,j), vArr(i,j)
      else
        row.Add CStr(j), vArr(i,j)
      end if
    Next
    Call c.Add(row)
  next

  set CreateFromArray2D = Create(name, stdEnumerator.CreateFromIEnumVariant(c), vHeaders)
End Function

'Create a table from an excel range
'@param rng as Object<Excel.Range> - The range to create the table from
'@param Optional headers as Boolean - If true, the first row of the range is used as the headers. If false, the headers are not set. If an array of strings, the headers are set to numbers.
'@returns - A new table containing the data from the range
Public Function CreateFromExcelRange(ByVal rng as Object, Optional ByVal headers as boolean = true) as stdTable
  set CreateFromExcelRange = CreateFromArray2D(rng.address, rng.Value, headers)
End Function

'Create a new dictionary from a recordset
'@protected
'@param rs as Object<ADODB.RecordSet> - The recordset to create the dictionary from
'@param lastItem as Variant<Object<ADODB.RecordSet>|null> - The last dict returned.
'@param index as long - The index of the item to be returned from the recordset. If this is the first call, pass 1.
'@returns Dictionary<string,variant> - A new dictionary containing the data from the recordset
Public Function protGetDictFromRecordSet(ByVal rs as Object, ByVal lastItem as variant, ByVal index as long) as Variant
  if index > 1 then Call rs.MoveNext()
  if rs.EOF then
    protGetDictFromRecordSet = Null
    Exit Function
  end if

  set protGetDictFromRecordSet = CreateObject("Scripting.Dictionary")
  protGetDictFromRecordSet.Add "__index", index
  Dim i as long
  For i = 0 to rs.fields.count - 1
    protGetDictFromRecordSet.Add rs.fields(i).Name, rs.fields(i).Value
  Next i
End Function

'Get a field of an object
'@protected
'@param field - The field to get
'@param dict as Object<Dictionary<string,variant>> - The object to get the field from
'@returns Variant - The value of the field
Public Function protGetDictionaryField(ByVal field as string, ByVal dict as Object, ByVal keyToIgnore as variant) as Variant
  Call CopyVariant(protGetDictionaryField, dict.item(field))
End Function

'Initialize the current table
'@param name as string - The name of the table
'@param headers as Variant - The headers of the table
'@param rows as stdEnumerator - The rows of the table
Public Sub protInit(name as string, headers as Variant, rows as stdEnumerator)
  With This
    .Name = name
    .Headers = headers
    Set .Rows = Rows
  End With
End Sub


'Get/Set the name of the table
Public Property Get Name() as string
  Name = This.Name
End Property
Public Property Let Name(ByVal value as string)
  This.Name = value
End Property

'Get the headers of the table
Public Property Get Headers() as Variant
  Headers = This.Headers
End Property

'Get the rows of the table
'@returns stdEnumerator<Dictionary<string,variant>> - An enumerator containing the rows of the table
Public Property Get Rows() as stdEnumerator
  Set Rows = This.Rows
End Property

'Apply a lambda to each row of the current table
'@param lambda as stdICallable<(row:Object)=>Void> - A lambda or callback that is applied to each row. This can be used to update the rows in place.
'@returns stdTable - The existing table, including any mutated rows
Public Function ForEach(ByVal lambda as stdICallable) as stdTable
  set Update = Me
  Call This.rows.foreach(lambda)
End Function

'Filter the current table using the specified lambda
'@param lambda as stdICallable<(row:Object)=>boolean> - A lambda or callback that takes a row and returns a boolean. If true, the row is kept.
'@returns stdTable - A new table containing only the rows that passed the filter
Public Function Filter(ByVal lambda as stdICallable) as stdTable
  set Filter = new stdTable
  Call Filter.protInit(This.name, This.Headers, This.Rows.Filter(lambda))
End Function

'Creates a new table containing only the specified fields from the current table
'@param Fields() as string - The fields to select
'@returns stdTable - A new table containing only the specified fields
Public Function FieldsSelect(ParamArray Fields()) as stdTable
  set FieldsSelect = Me
  Redim Preserve Fields(1 to UBound(Fields)+1)
  This.Headers = Fields
End Function

'Remove fields from the current table
'@param ParamArray<> - The fields to remove
'@returns stdTable - The existing table with the fields removed
Public Function FieldsRemove(ParamArray fields()) as stdTable
  set FieldsRemove = Me
  Dim toRemove as Object: set toRemove = CreateObject("Scripting.Dictionary")
  Dim i as long
  For i = 0 to UBound(fields)
    toRemove.Add fields(i), true
  Next i

  'Remove from headers
  Dim vNewHeaders() as string: Redim vNewHeaders(1 to UBound(This.Headers) - UBound(fields))
  Dim index as long: index = 1
  For i = 1 to UBound(This.Headers)
    if not toRemove(This.Headers(i)) then
      vNewHeaders(index) = This.Headers(i)
      index = index + 1
    end if
  next
  This.Headers = vNewHeaders
End Function

'Update a field of the current table using the specified lambda
'@param field - The name of the field to update
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to set the field to
'@returns stdTable - The existing table.
Public Function FieldUpdate(ByVal field as string, ByVal lambda as stdICallable) as stdTable
  set FieldUpdate = Me
  Dim row as object
  For each row in This.Rows
    row(field) = lambda.Run(row)
  next
End Function

'Rename fields
'@param Paramarray<> - The fields to rename. Each pair of arguments is a field to rename and the new name for that field.
'@returns stdTable - The existing table with mutated field names
Public Function FieldsRename(ParamArray fields()) as stdTable
  set FieldsRename = Me

  Dim i as long
  For i = 0 to UBound(fields) Step 2
    Dim oldName as string: oldName = fields(i)
    Dim newName as string: newName = fields(i+1)
    Dim j as long, renamed as boolean: renamed = false
    For j = 1 to UBound(This.Headers)
      if This.Headers(j) = oldName then 
        This.Headers(j) = newName
        Renamed = true
        Exit For
      end if
    Next j
    if not renamed then Err.Raise 5, "stdTable.FieldsRename", "Field '" & oldName & "' not found"
  Next i

  Dim row as object
  For each row in this.rows
    For i = 0 to Ubound(fields) step 2
      row.item(fields(i+1)) = row.item(fields(i))
      Call row.Remove(fields(i))
    Next i
  next row
End Function

'Add a field to the current table
'@param field as string - The name of the field to add
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to set the field to
'@returns stdTable - The existing table with the new field
Public Function FieldAdd(ByVal field as string, ByVal lambda as stdICallable) as stdTable
  set FieldAdd = Me
  
  'Ensure field doesn't already exist
  Dim i as long: For i = 1 to ubound(this.headers)
    if this.headers(i) = field then Err.Raise 5, "stdTable.FieldAdd", "Field '" & field & "' already exists"
  next

  'Add to fields
  Redim Preserve This.Headers(1 to UBound(This.Headers)+1)
  This.Headers(UBound(This.Headers)) = field

  'Map lambda to field
  Dim row as object
  For each row in This.Rows
    row(field) = lambda.Run(row)
  next
End Function

'Expand object fields
'@param field as string - The field to expand
'@returns stdTable - The existing table.
'@example - For the table `[{"Car":{"Doors":2}}, {"Car":{"Doors":3}}, ...]`
'           The code `FieldExpand("Car")` would return the table `[{"Doors":2,"Car":{"Doors":2}}, {"Doors":3,"Car":{"Doors":2}}, ...]`
Public Function FieldExpand(ByVal field as string) as stdTable
  set FieldExpand = Me
  For each row in This.Rows
    if isObject(row(field)) then
      Dim target as Object: set target = row(field)
      select case typename(target)
        case "Dictionary"
          Dim key as variant
          For each key in target.keys()
            Call CopyVariant(row(key), target(key))
          next
        case else
          Dim prop as variant
          For each prop in stdCOM.Create(target).Properties
            Call CopyVariant(row(prop), stdCallback.CreateFromObjectProperty(target, prop, vbGet).Run())
          next
      end select
    else
      Err.Raise 5, "stdTable.FieldExpand", "Field '" & field & "' does not contain an object. Cannot expand."
    end if
  next
End Function

'Update a field of the current table to a specified value
'@param field - The name of the field to update
'@returns stdTable - The existing table.
Public Function FieldUpdateStatic(ByVal field as string, ByVal value as variant) as stdTable
  set FieldUpdateStatic = Me
  Dim row as object
  For each row in This.Rows
    row(field) = value
  next
End Function

'Link 1 table to another 
'@param table2 as stdTable - The table to join to
'@param field1 - The field in the current table to join on
'@param field2 - The field in the other table to join on
'@param joinAlias - The alias to in the join field
'@param returnOne - If true, only return the first matching row from the other table. Otherwise, return all matching rows. If no rows match, return Null if returnOne is true, else return an empty stdEnumerator.
'@returns stdTable - The existing table with mutated join values
Public Function Join(ByVal table2 as stdTable, ByVal field1 as string, ByVal field2 as string, Optional ByVal joinAlias as string = "", Optional ByVal returnOne as Boolean = false) as stdTable
  set Join = Me
  Dim groups as Object: set groups = table2.Rows.groupBy(stdCallback.CreateFromObjectMethod(Me, "protGetDictionaryField").bind(field2))
  if returnOne then
    Dim vKey: For each vKey in groups.keys()
      groups(vKey) = groups(vKey).item(1)
    next
  end if
  For each obj in This.Rows
    if obj("join") is nothing then set obj("join") = CreateObject("Scripting.Dictionary")
    if joinAlias = "" then joinAlias = table2.Name
    Dim sKey as string: sKey = obj(field1)
    if groups.exists(sKey) then
      set obj("join")(joinAlias) = groups(sKey)
    else
      if returnOne then
        obj("join")(joinAlias) = Null
      else
        obj("join")(joinAlias) = stdEnumerator.CreateEmpty()
      end if
    end if
  next
End Function 

'Concat rows from a Collection to the current table
'@param c as Collection - The collection to concat
'@returns stdTable - The existing table with the new rows
'@remark - The collection must contain objects with the same fields as the current table. AddRow() can be used to add a single row but is significantly slower than using a collection.
Public Function Concat(ByVal c as Collection) as stdTable
  set Concat = new stdTable
  Call Concat.protInit(this.name, this.headers, this.rows.concat(c))
End Function

'Add a row to the current table
'@param ParamArray<Variant> - The values to add to the row. The number of values must match the number of fields in the table.
'@returns stdTable - The existing table with the new row
'@remark - Use Concat() to add multiple rows. This method will be significantly slower.
Public Function AddRow(ParamArray values()) as stdTable
  set AddRow = Me
  if UBound(values) <> UBound(This.Headers) then Err.Raise 5, "stdTable.AddRow", "Number of parameters must match number of fields"
  Dim row as object: set row = CreateObject("Scripting.Dictionary")
  Dim i as long: For i = 1 to UBound(This.Headers)
    row.Add This.Headers(i), values(i)
  next
  Dim c as new Collection: Call c.Add(row)
  set AddRow = Concat(c)
End Function

'Reverse the order of the rows in the current table
'@returns stdTable - The existing table with the rows reversed
Public Function Reverse() as stdTable
  set Reverse = new stdTable
  Call Reverse.protInit(This.name, This.Headers, This.Rows.Reverse())
End Function

'Group the current table by the specified field. Each row generated by this function will have 2 fields, 1 real and 1 virtual. The virtual field will not be part of the object but not in the headers (and not export in any To___() methods).
'Output is of the form: `[{key:..., group:stdEnumerator<Dictionary<string,variant>>},...]`
'@param fieldName as string - The name of the field to group by
'@param virtualName as string - The name of the virtual field to store the grouped rows in
'@returns - A new table containing the grouped rows
'@remark - This function uses `GroupBy` internally.
Public Function GroupByField(ByVal fieldName as string, Optional ByVal virtualName as string = "group") as stdTable
  set GroupByField = GroupBy(stdCallback.CreateFromObjectMethod(Me, "protGetDictionaryField").bind(fieldName), fieldName, virtualName)
End Function

'Group the current table by the specified lambda. Each row generated by this function will have 2 fields, 1 real and 1 virtual. The virtual field will not be part of the object but not in the headers (and not export in any To___() methods).
'Output is of the form: `[{key:..., group:stdEnumerator<Dictionary<string,variant>>},...]`
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to group by
'@param keyName as string - The name of the field to store the key in
'@param virtualName as string - The name of the virtual field to store the grouped rows in
'@returns - A new table containing the grouped rows
Public Function GroupBy(ByVal lambda as stdICallable, Optional ByVal keyName as string = "groupKey", Optional ByVal virtualName as string = "group") as stdTable
  set GroupBy = new stdTable
  set groups = This.Rows.GroupBy(lambda)
  Dim cRows as Collection: set cRows = new Collection
  Dim vKey: For each vKey in groups.keys()
    Dim newRow as object: set newRow = CreateObject("Scripting.Dictionary")
    newRow.Add keyName, vKey
    newRow.Add virtualName, groups(vKey)
    cRows.Add newRow
  next
  Dim headers() as String: Redim headers(1 to 1)
  headers(1) = keyName
  Call GroupBy.protInit(This.name, headers, stdEnumerator.CreateFromIEnumVariant(cRows))
End Function

'Deep copy the current table
'@returns stdTable - A new table with the same data as the current table
Public Function Clone() as stdTable
  set Clone = new stdTable
  Dim cRows as Collection: set cRows = new Collection
  Dim row as Object
  For each row in This.Rows
    Dim newRow as object: set newRow = CreateObject("Scripting.Dictionary")
    Dim key as variant
    For each key in This.Headers
      newRow.Add key, row(key)
    next
    newRow.Add "join", row("join") 'maintain joins
    cRows.Add newRow
  next
  Call Clone.protInit(This.name, This.Headers, stdEnumerator.CreateFromIEnumVariant(cRows))
End Function

'Remove duplicates based on some generated id
'@param lambda as stdICallable<(row:Object)=>variant> - A lambda or callback that takes a row and returns a value to group by to find unique values
'@returns stdTable - A new table with the duplicates removed
Public Function Unique(ByVal lambda as stdICallable) as stdTable
  set Unique = new stdTable
  Call Unique.protInit(This.name, This.Headers, This.Rows.Unique(lambda))
End Function

'Export the table to a 2D variant array
'@returns Variant - A 2D variant array containing the data from the table
Public Function ToArray2D() as Variant
  Dim arr() as Variant: Redim arr(1 to This.Rows.Length + 1, 1 to UBound(This.Headers))
  Dim j as long
  For j = 1 to UBound(This.Headers)
    arr(1,j) = This.Headers(j)
  Next j
  Dim i as long: i = 1
  For each row in this.rows
    i = i + 1
    For j = 1 to UBound(This.Headers)
      arr(i,j) = row(This.Headers(j))
    Next j
  Next row
  ToArray2D = arr
End Function

'Export the table to a range
'@param rng as Range - The range to export the table to
'@returns Range - The full range that was created
Public Function ToExcelRange(ByVal rng as Object) as Object
  Dim arr as Variant: arr = ToArray2D()
  Dim numRows as Long: numRows = UBound(arr,1)
  Dim numCols as Long: numCols = UBound(arr,2)
  set ToExcelRange = rng.Resize(numRows, numCols)
  ToExcelRange.Value2 = arr
End Function

'Export the table to a list object
'@param rng as Range - The range to export the table to
'@param Optional tableName as string - The name of the table. If not specified, this table's name is used.
'@param Optional insertRows as boolean - If true, insert rows into the worksheet. Otherwise, overwrite the existing range.
'@returns ListObject - The list object that was created
'@remark - If a list object already exists in the specified range, it will be overwritten.
Public Function ToListObject(ByVal rng as Range, Optional ByVal tableName as string = "", Optional ByVal insertRows as boolean = false, optional ByVal useCurrentRegion as Boolean = true) as ListObject
  Dim arr() as Variant: arr = ToDataBodyArray()
  Dim iUB1 as Long: iUB1 = UBound(arr,1)+1 '+1 for headers
  Dim iUB2 as Long: iUB2 = UBound(arr,2)
  Dim ws as Worksheet: set ws = rng.Worksheet
  Dim lo as ListObject: set lo = rng.ListObject
  if not lo is nothing then set rng = lo.range.resize(1,1)
  if tableName = "" then 
    if lo is nothing then 
      tableName = This.Name
    else
      tableName = lo.Name
    end if
  end if

  'Calculate original range
  Dim originalRange as Object
  if not lo is nothing then
    set originalRange = lo.range
  elseif useCurrentRegion and (rng.currentRegion.rows.count > 1 or rng.currentRegion.columns.count > 1) then
    set originalRange = rng.currentRegion
  else
    set originalRange = rng
  end if

  'Insert rows if needed
  if insertRows then
    if originalRange.rows.count < iUB1 then Call originalRange.offset(originalRange.rows.count).resize(originalRange.rows.count - iUB1).insert(xlShiftDown)
    if originalRange.columns.count < iUB2 then Call originalRange.offset(,originalRange.columns.count).resize(,originalRange.columns.count - iUB2).insert(xlShiftRight)
  end if

  'resize original range and existing listobject to match array
  Dim hRng as Range: set hRng = originalRange.resize(1, iUB2)
  Dim dRng as Range: set dRng = originalRange.offset(1).resize(iUB1-1, iUB2)
  Dim loRng as Range: set loRng = originalRange.resize(iUB1, iUB2)
  if not lo is nothing then Call lo.resize(loRng)
  hRng.value2 = This.Headers
  dRng.value2 = arr
  
  'create listobject if needed
  if lo is nothing then 
    set ToListObject = loRng.parent.ListObjects.Add(xlSrcRange, loRng, xlListObjectHasHeaders:=xlYes)
    ToListObject.Name = tableName
  else
    set ToListObject = lo
  end if
End Function

'Export the table to a recordset
'@returns Object<ADODB.RecordSet> - An ADODB recordset containing the data from the table
Public Function toRecordSet() as Object
  Const adLockPessimistic As Long = 2
  Const adVariant As Long = 12
  Dim rs as Object: set rs = CreateObject("ADODB.RecordSet")
  Dim i as long
  rs.Fields.Append "_ID", adInteger
  For i = 1 to UBound(This.Headers)
    rs.Fields.Append This.Headers(i), adVariant
  Next i
  rs.LockType = adLockPessimistic
  rs.Open
  Dim row as Object
  i=0
  For each row in This.Rows
    rs.AddNew
    i=i+1
    rs.Fields("ID").Value = i
    For i = 1 to UBound(This.Headers)
      rs.Fields(This.Headers(i)).Value = row(This.Headers(i))
      rs.update
    Next i
  Next
  set toRecordSet = rs
End Function

'Export the table to a collection of dictionaries
'@returns Collection<Dictionary<string,variant>> - A collection of dictionaries containing the data from the table
'@remark - Strips virtual fields
Public Function ToCollection() as Collection
  Dim c as Collection: set c = new Collection
  Dim row as Object
  For each row in This.Rows
    Dim rowNew as Object: set rowNew = CreateObject("Scripting.Dictionary")
    Dim vHeader
    For each vHeader in This.Headers
      rowNew.Add vHeader, row(vHeader)
    Next
    c.Add row
  Next
  Set ToCollection = c
End Function

'Export table to JSON
'@returns - A JSON string containing the data from the table
'@requires stdJSON
Public Function ToJSON() as string
  With stdJSON.Create(eJSONArray)
    Dim row as Object
    For each row in This.Rows.ToCollection()
      Call .Add(row)
    Next
    ToJSON = .ToString()
  End With
End Function







'Copy a variant from src to destination
'@param dest - The destination variant
'@param src - The source variant
Private Sub CopyVariant(ByRef dest as variant, ByVal src as variant)
  if isObject(src) then
    set dest = src
  else
    let dest = src
  end if
End Sub

'Export the table to a 2D variant array
'@returns Variant - A 2D variant array containing the data from the table
Private Function ToDataBodyArray() as Variant
  Dim arr() as Variant: Redim arr(1 to This.Rows.Length, 1 to UBound(This.Headers))
  Dim i as long: i = 0
  For each row in this.rows
    i = i + 1
    Dim j as long
    For j = 1 to UBound(This.Headers)
      arr(i,j) = row(This.Headers(j))
    Next j
  Next row
  ToDataBodyArray = arr
End Function