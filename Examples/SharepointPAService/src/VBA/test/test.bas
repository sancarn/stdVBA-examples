Attribute VB_Name = "SPPAServiceTest"

Sub main()
    Dim service as SPPAService: set service = SPPAService.CreateFromCombinedConfig("C:\Temp\SPPAService_NOS.json")
    set items = service.GetListItems().Await().ResponseMapped()

End Sub