Attribute VB_Name = "stdTimerTests"
Public Sub Test()
    Call StartTimer(1000, stdCallback.CreateFromModule("stdTimerTests", "TestCallback"))
End Sub
Public Sub TestCallback()
    Static i As Long: i = i + 1
    Debug.Print "hello " & i
End Sub
