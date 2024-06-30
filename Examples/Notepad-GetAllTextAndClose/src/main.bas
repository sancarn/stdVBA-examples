Attribute VB_Name = "main"
Sub executeMain()
  Call execute(True)
End Sub
Sub executeNoQuit()
  Call execute(False)
End Sub

Sub execute(ByVal quitAfter As Boolean)
  With Sheet1
    .UsedRange.Clear
    .[a1] = "WindowTitle"
    .[b1] = "Value"
    Dim r As Range: Set r = .Range("A1")
    
    Dim iOffset As Long: iOffset = 0
    Dim procs As Collection: Set procs = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.Name like ""*Notepad.exe*"""))
    Dim proc As stdProcess
    For Each proc In procs
      Dim wnd As stdWindow
      For Each wnd In stdWindow.CreateManyFromProcessId(proc.id)
        Dim main As stdAcc: Set main = stdAcc.CreateFromHwnd(wnd.handle)
        Dim title As String: title = main.name
        
        'Activate all tabs to populate window accessible controls with text
        Dim accTab As stdAcc
        For Each accTab In main.FindAll(stdLambda.Create("$1.Role = ""ROLE_PAGETAB"""))
          Call accTab.DoDefaultAction
        Next
        
        Dim editWnd As stdWindow
        For Each editWnd In wnd.FindAll(stdLambda.Create("$1.Class = ""RichEditD2DPT"""))
          iOffset = iOffset + 1
          
          Dim editAcc As stdAcc: Set editAcc = stdAcc.CreateFromHwnd(editWnd.handle).children(4)
          Dim text As String: text = editAcc.value
          
          r.offset(iOffset, 0) = title
          r.offset(iOffset, 1).Value2 = "'" & text
        Next
      Next
      
      If quitAfter Then Call proc.ForceQuit(400)
    Next
  End With
End Sub
