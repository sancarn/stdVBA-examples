Attribute VB_Name = "SPPAServiceTest"
Sub testListItems()
    Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_NOS.json")
    Dim j As stdJSON: Set j = service.ListItems().Await().ResponseMapped()
    Call JSONViewer.ShowViewer(j)
End Sub

Sub testListItemsWithQuery()
    Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_NOS.json")
    Dim j As stdJSON: Set j = service.ListItems("$expand=NOS&$select=Id,County,NOS/Title").Await().ResponseMapped()
    Call JSONViewer.ShowViewer(j)
End Sub

Sub testListItem()
  Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_NOS.json")
  Dim j As stdJSON: Set j = service.ListItem(1).Await().ResponseMapped()
  Call JSONViewer.ShowViewer(j)
End Sub

Sub testGetListItemType()
  Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_TacticalData-Test.json")
  Debug.Print service.getListItemType().Await().ResponseMapped()
End Sub

Sub testListItemCreate()
  Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_TacticalData-Test.json")
  Dim data As stdJSON: Set data = stdJSON.Create()
  Call data.Add("Title", "Hello")
  Call data.Add("Poop", "Choice 1")
  Dim HTTP As stdHTTP: Set HTTP = service.ListItemCreate(data).Await()
  Debug.Print HTTP.ResponseStatus
End Sub

Sub testListItemCreateBulk()
  Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_TacticalData-Test.json")
  Dim items As stdArray: Set items = stdArray.Create()
  For i = 1 To 10
    With stdJSON.Create()
      .Add "Title", "Hello " & i
      .Add "Poop", "Choice 1"
      Call items.Push(.ToSelf())
    End With
  Next
  Debug.Print "Site: " & service.SiteURL & vbCrLf & "List:  " & service.ListSelector
  
  With service.ListItemsCreateBatch(items)
    Dim results as string: results = .map(stdLambda.Create("$1.Await().ResponseText")).join(vbCrLf & vbCrLf)
  End with
End Sub

Sub testListItemsDeleteBatch()
  Dim service As SPPAService: Set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_TacticalData-Test.json")
  Dim items As stdArray: Set items = stdArray.Create(31,32,33)
  Debug.Print "Site: " & service.SiteURL & vbCrLf & "List:  " & service.ListSelector
  
  With service.ListItemsDeleteBatch(items)
    Dim results as string: results = .map(stdLambda.Create("$1.Await().ResponseText")).join(vbCrLf & vbCrLf)
  End with
End Sub


Sub t()
  Debug.Print stdJSON.CreateFromParams(eJSONObject, "a", 1, "b", 2).ToString()
  Debug.Print stdJSON.CreateFromParams(eJSONArray, "a", 1, "b", 2).ToString()
End Sub
