# xlTableTools

Helper functions which can be used in combination with Excel `ListObject` tables.

## Setup test

You can run this setup in order to test the examples in this document.

```vb
Sub RunSetup()
  Range("A1:A10").value = Application.Transpose(Array("ID", 1, 2, 3, 4, 5, 6, 7, 8, 9))
  Range("B1:D1").value = Array("Score", "Category", "CategoryName")
  With ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:D10"), , xlYes)
      .name = "tblMain"
  End With
      
  Range("J1:J4").value = Application.Transpose(Array("ID", "H", "M", "L"))
  Range("K1:K4").value = Application.Transpose(Array("Name", "High", "Medium", "Low"))
  With ActiveSheet.ListObjects.Add(xlSrcRange, Range("J1:K4"), , xlYes)
      .name = "tblCategories"
  End With
End Sub
```

## Example

The following example will assign values to `Score`, `Category` and `CategoryName` fields of the table. Then if there are any `Category` Highsm, the table will filter to these.

```vb
Sub Test()
  Dim loCats As ListObject: Set loCats = ActiveSheet.ListObjects("tblCategories")
  Dim rels As Collection: Set rels = New Collection
  rels.Add CreateRelationship("Category", "Category", loCats, "ID")
  
  Dim lo As ListObject: Set lo = ActiveSheet.ListObjects("tblMain")
  lo.AutoFilter.ShowAllData

  Call updateField(lo, "Score", stdLambda.Create("rnd(1)"))
  Call updateField(lo, "Category", stdLambda.CreateMultiline(Array( _
    "if $1.Score < 0.3 then ""L""", _
    "else if $1.Score < 0.7 then ""M""", _
    "else ""H""", _
    "end" _
  )))
  Call updateField(lo, "CategoryName", stdLambda.Create("$1.rels__.Category.Name"), Relationships:=rels)

  Stop

  Call applyFilterToTable(lo, "ID", stdLambda.Create("$1.Category = ""H"""))
End Sub
```

## Functions

* `updateFieldConst` - Update a field of a table in Excel to a value.
* `updateField` - Update a field of a table by callback in Excel.
* `CreateRelationship` - Create a relationship object. Links some table's to the `ToTable` where some fields match.
* `applyFilterToTable` - Applies a filter to a table by callback in Excel. The table requires an ID field.