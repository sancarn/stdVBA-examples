Attribute VB_Name = "ConditionalFormattingEx"
Public Sub BulkApplyFormatting()
  Call TargetedApplyFormatting(shTest.UsedRange, True)
End Sub

Public Sub TargetedApplyFormatting(ByVal Target As Range, Optional ByVal forceRefreshStyles = False)
  Static eConditions As stdEnumerator
  If eConditions Is Nothing Or forceRefreshStyles Then
    Set eConditions = stdEnumerator.CreateFromListObject(shLookups.ListObjects("ConditionalFormatting"))
    Call eConditions.ForEach(stdLambda.Create("set $2.lambda = $1.Create($2.Lambda): $2").Bind(stdLambda))
  End If
  
  Dim cell As Range
  For Each cell In Target.Cells
    If cell.Value <> Empty Then
      Dim row As Object
      Set row = eConditions.FindFirst(stdLambda.Create("$2.lambda.run($1)").Bind(cell.Value), Nothing)
      If Not row Is Nothing Then
        Dim colors: colors = Split(row("InteriorColor"), ",")
        cell.Interior.Color = RGB(colors(0), colors(1), colors(2))
      End If
    Else
      cell.Interior.ColorIndex = xlColorIndexNone
    End If
  Next
End Sub


