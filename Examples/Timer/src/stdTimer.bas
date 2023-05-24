Attribute VB_Name = "stdTimer"
#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
    Private Declare PtrSafe Function KillTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long) As Long
#Else
    Enum LongPtr
        [_]
    End Enum
    Private Declare Function SetTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
    Private Declare Function KillTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long) As Long
#End If

Private Type TimerInfo
    isAlive As Boolean
    id As Long
    callable As stdICallable
    times As Long
End Type
Private Timers() As TimerInfo
Private pTimerCount As Long

'Specify a time in milliseconds and a function, and the routine will call this function every n milliseconds.
'@param {Long} The frequency at which the function should be called
'@param {stdICallable} The function to call at the frequency specified
'@param {Long} The number of times to call the callback
'@returns {Long} Index of timer in Timers array. Use this when calling StopTimer()
Public Function StartTimer(ByVal iMilliseconds As Long, ByVal callable As stdICallable, Optional ByVal iNumberOfTimes As Long = -1) As Long
    On Error Resume Next: Dim iNextTimer As Long: iNextTimer = UBound(Timers) + 1: On Error GoTo 0
    ReDim Timers(0 To iNextTimer)
    pTimerCount = pTimerCount + 1
    With Timers(iNextTimer)
        .isAlive = True
        Set .callable = callable
        .times = iNumberOfTimes
        .id = SetTimer(0, 0, iMilliseconds, AddressOf OnTime)
        StartTimer = iNextTimer
    End With
End Function

'Stop an active timer
'@param {Long} The index of the timer to stop in Timers array. This ID is returned from StartTimer()
Public Sub StopTimer(ByVal iTimerIndex As Long)
    With Timers(iTimerIndex)
        If .isAlive Then
            Call KillTimer(0, .id)
            .isAlive = False
            pTimerCount = pTimerCount - 1
        End If
    End With
End Sub

'Stop all timers
Public Sub StopAll()
    For i = 0 To UBound(Timers)
        With Timers(i)
            Call KillTimer(0, .id)
            .isAlive = False
        End With
    Next
    pTimerCount = 0
End Sub

'Callback called at the desired frequency
Private Sub OnTime(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Integer)
    'If VBA reset, make sure to kill the timer
    If pTimerCount = 0 Then
        Call KillTimer(0, idEvent)
        Exit Sub
    End If
    
    'Get timer and run callback
    Dim iTimerIndex As Long: iTimerIndex = getTimerIndex(idEvent)
    With Timers(iTimerIndex)
        Call .callable.Run
        
        'Handle callback number of times
        If .times > 0 Then
            .times = .times - 1
            If .times = 0 Then Call StopTimer(iTimerIndex)
        End If
    End With
End Sub

'Obtain the index of a timer from it's internal ID
Private Function getTimerIndex(ByVal nID As Long) As Long
    For i = 0 To UBound(Timers)
        With Timers(i)
            If .id = nID Then
                getTimerIndex = i
                Exit Function
            End If
        End With
    Next
End Function
