Attribute VB_Name = "dsMain"
Sub DataSources_Refresh()
  Dim Sources: Sources = Array(dsSrcExcelRange, dsSrcPowerBI, dsSrcPowerQuery)
  
  Dim lo As ListObject: Set lo = shDataSources.ListObjects("DataSources")
  Dim fibers As Collection: Set fibers = New Collection
  
  Dim lr As ListRow
  For Each lr In lo.ListRows
    Dim TableName As String: TableName = getValueRangeByHeader(lr, "Name").value
    Dim OutputPath As String: OutputPath = getValueRangeByHeader(lr, "OutputPath").value
    Dim SourceType As String: SourceType = getValueRangeByHeader(lr, "Type").value
    Dim Data As stdJSON: Set Data = stdJSON.CreateFromString(getValueRangeByHeader(lr, "Data").value)
    
    'Create fiberTemplate ready for processing
    Dim fiberTemplate As stdFiber: Set fiberTemplate = stdFiber.Create(SourceType & " " & TableName)
    With fiberTemplate
      Set .Meta("Row") = lr
      .Meta("Table") = TableName
    End With
    
    Dim frequency As String: frequency = getValueRangeByHeader(lr, "Frequency").value
    Dim lastUpdated As Date: lastUpdated = getValueRangeByHeader(lr, "Out-DateExtracted").value
    Dim mainFiber As stdFiber: Set mainFiber = Nothing
    If checkNeedsRefresh(frequency, lastUpdated) Then
      Dim Source As dsISrc
      For Each Source In Sources
        If Source.getName() = SourceType Then
          Set mainFiber = Source.linkFiber(fiberTemplate, OutputPath, Data)
        End If
      Next
      
      If Not mainFiber Is Nothing Then
        Call mainFiber.add(stdCallback.CreateFromModule("dsMain", "ProcessCleanup"), "X. Cleanup")
        Call mainFiber.addStepChangeHandler(stdCallback.CreateFromModule("dsMain", "ProcessStep"))
        Call mainFiber.addErrorHandler(stdCallback.CreateFromModule("dsMain", "ProcessCleanup"))
        Call fibers.add(mainFiber)
      End If
    End If
  Next
  
  If fibers.count > 0 Then
    Dim agentInit As stdICallable: Set agentInit = stdCallback.CreateFromModule("dsMain", "AgentInit")
    Dim agentDestroy As stdICallable: Set agentDestroy = stdCallback.CreateFromModule("dsMain", "agentDestroy")
    Dim runningCB As stdICallable: Set runningCB = stdCallback.CreateFromModule("dsMain", "RunningCallback")
    Call stdFiber.runFibers(fibers, 4, agentInit, agentDestroy, runningCB)
    Application.StatusBar = Empty
    MsgBox "All datasets downloaded", vbInformation
  Else
    MsgBox "No datasets required", vbInformation
  End If
End Sub

'Initialise agent callback
'@param agent as Object<Scripting.Dictionary> - The agent object created
'@remark Current use is creating Excel.Application instances, as many fibers (ExcelFile, Sharepoint, PowerBI) require Excel instances, and this can keep the process asynchronous
Public Sub agentInit(ByVal Agent As Object)
  Dim xlApp As Excel.Application: Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  xlApp.WindowState = xlMinimized
  Set Agent("xl") = xlApp
End Sub


Public Sub agentDestroy(ByVal Agent As Object)
  Dim xlApp As Excel.Application: Set xlApp = Agent("xl")
  xlApp.DisplayAlerts = False
  Call xlApp.Quit
  Set Agent("xl") = Nothing
End Sub

'Called every time an update is triggered, indicates the current progress of the process
Public Sub RunningCallback(ByVal iFinishedCount As Long, iFiberCount As Long)
  Static symbols As Variant: If isEmpty(symbols) Then symbols = Array("/", "-", "\", "|")
  Static index As Long: index = index + 1
  If index = 5 Then index = 1
  Application.StatusBar = iFinishedCount & "/" & iFiberCount & " processing " & symbols(index - 1)
End Sub

'Final stage of the process, update Out-DataExtracted,
'@fiberRunner
'@protected
Public Function ProcessCleanup(ByVal fiber As stdFiber) As Boolean
  Dim lr As ListRow: Set lr = fiber.Meta("Row")
  If fiber.errorText = "" Then
    getValueRangeByHeader(lr, "Out-DateExtracted").value = Now()
    getValueRangeByHeader(lr, "Out-ErrorText").value = ""
    getValueRangeByHeader(lr, "Out-Step").value = "Complete"
  Else
    getValueRangeByHeader(lr, "Out-ErrorText").value = fiber.StepName & " " & fiber.errorText
    getValueRangeByHeader(lr, "Out-Step").value = "Error"
  End If
  ProcessCleanup = True
End Function

'Updates Out-Step on step changes
'@fiberStepHandler
'@protected
Public Sub processStep(ByVal fiber As stdFiber)
  getValueRangeByHeader(fiber.Meta("Row"), "Out-Step").value = fiber.StepName
End Sub

'Obtain a range for a listrow based on it's header
'@param row - The list row to obtain the data for
'@param header - The header to obtain the cell for.
'@returns - The cell represented by the header for the specified list row
Private Function getValueRangeByHeader(ByVal row As ListRow, ByVal header As String) As Range
  Dim lo As ListObject: Set lo = row.Parent
  Set getValueRangeByHeader = Application.Intersect(row.Range, lo.ListColumns(header).Range)
End Function

'Checks whether a refresh is required given a frequency and previous refresh date
'@param frequency - String representing frequency refresh is required at. E.G. "Monthly", "Quarterly", "Annually"
'@param oldRefreshDate - The date the last refresh of the dataset was committed and complete successfully
'@returns - True if refresh required, false otherwise.
Private Function checkNeedsRefresh(ByVal frequency As String, ByVal oldRefreshDate As Date) As Boolean
  Select Case frequency
    Case "Monthly"
      checkNeedsRefresh = (Now() - oldRefreshDate) > 30
    Case "Quarterly"
      checkNeedsRefresh = (Now() - oldRefreshDate) > 90
    Case "Annually"
      checkNeedsRefresh = (Now() - oldRefreshDate) > 365
  End Select
End Function




