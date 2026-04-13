Attribute VB_Name = "mMain"

Sub mainRunUpdater()
  Dim row As ORMMissingDataRow
  Set row = ORMMissingDataRow.CreateUpdator("BIG_TEST_1")
  
  row.AssetType = EAMT_NotApplicable
  row.Division = EORMDivision_NotApplicable
  row.ORMFunction = EORMFunction_NotApplicable
  
  Dim c As New Collection
  c.Add row
  
  Call ORMMissingDataRow.ExecuteCollection(c)
End Sub
