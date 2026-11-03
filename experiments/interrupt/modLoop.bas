Attribute VB_Name = "modLoop"
' @desc CPU-bound loop logging a tick each second for interrupt probe
Option Explicit

' @desc Busy-waits 60 seconds and logs a tick record every second
Public Sub LongLoop()
    Dim t0 As Double, last As Double, n As Long
    t0 = Timer: last = t0
    Do While Timer - t0 < 60
        If Timer - last >= 1 Then
            last = Timer: n = n + 1
            LogShim.LogJ "tick", "{""n"":" & n & "}"
            modCancel.Check
        End If
    Loop
    LogShim.LogJ "loop_done", "{""n"":" & n & "}"
End Sub
