Attribute VB_Name = "SPPAServiceTest"

Sub main()
    Dim service as SPPAService: set service = SPPAService.CreateFromConfigFile("C:\Temp\SPPAService_NOS.json")
    set items = service.GetListItems("NOS County Owner")
End Sub