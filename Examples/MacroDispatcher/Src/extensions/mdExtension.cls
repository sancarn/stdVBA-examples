VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mdExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'TODO: Call various handles in process
'TODO: Consider how to best filter which extensions to use.
'TODO: Consider how to best handle the application object


Public Function Create() as mdExtension
  set Create = new mdExtension
  Call Create.protInit
End Function
Public Sub protInit()

End Sub

Public Function Handles() as Object
  set Handles = CreateObject("Scripting.Dictionary")
  Handles("Before_ApplicationCreate") = false
  Handles("After_ApplicationCreate") = false
  Handles("After_WorkbookOpen") = false
  Handles("On_Initialising") = false
  Handles("On_Ready") = false
  Handles("On_Executed") = false
  Handles("On_Running") = false
  Handles("On_Error") = false
  Handles("On_Completion") = false
End Function

'Make any changes necessary before the application object is created
'@param app - Passed by-ref to allow you to override the application object to be used
'@remark - Usually this method is used to set the application object to be used.
Public Sub Before_ApplicationCreate(ByRef app as Application)

End Sub

'Make any changes necessary after the application object has been created
'@param app - Application to make changes to if necessary
Public Sub After_ApplicationCreate(ByRef app as Application)

End Sub

'Called after a workbook has been opened
'@param wb - The workbook that has been opened, passed by-ref to allow you to make changes to the workbook if necessary
Public Sub After_WorkbookOpen(ByRef wb as Workbook)

End Sub

'Called when the mdJob is initialising
Public Sub On_Initialising()

End Sub

'Called when the mdJob is ready to be executed
Public Sub On_Ready()

End Sub

'Called when the mdJob is executed
Public Sub On_Executed()

End Sub

'Called when the mdJob is running
Public Sub On_Running()

End Sub

'Called when an error occurs
Public Sub On_Error(ByVal status as string)

End Sub

'Called when the mdJob has completed
'@param app - The application object that was used. If any clean up is required, this is the place to do it.
Public Sub On_Completion(ByVal app as Application)

End Sub