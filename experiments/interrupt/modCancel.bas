Attribute VB_Name = "modCancel"
' @desc Cancel flag check that stops all VBA via End for interrupt probe
Option Explicit

Private lastCheck As Double

' @desc Ends all running VBA when the cancel flag file exists (checked at most every 0.5 s)
Public Sub Check()
    If Abs(Timer - lastCheck) < 0.5 Then Exit Sub
    lastCheck = Timer
    If Len(Dir(ThisWorkbook.Path & "\_dbg\cancel.flag")) > 0 Then
        LogShim.LogJ "cancel_seen", "null"
        End
    End If
End Sub
