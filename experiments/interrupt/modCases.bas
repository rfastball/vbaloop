Attribute VB_Name = "modCases"
' @desc Window layout cases held busy for Win32 EXCEL7 title probing
Option Explicit

Private Sub Hold(label As String, Optional secs As Double = 12)
    Dim t0 As Double, last As Double
    LogShim.LogJ "case_ready", "{""case"":""" & label & """}"
    t0 = Timer: last = t0
    Do While Timer - t0 < secs
        If Timer - last >= 1 Then last = Timer: LogShim.LogJ "tick", "null"
    Loop
End Sub

' @desc Opens a second new workbook and holds Excel busy
Public Sub Case_SecondBook()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Hold "second"
    wb.Close SaveChanges:=False
End Sub

' @desc Opens a second workbook with its window hidden and holds Excel busy
Public Sub Case_HiddenBook()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Windows(1).Visible = False
    Hold "hidden"
    wb.Close SaveChanges:=False
End Sub

' @desc Opens a second window on the session workbook and holds Excel busy
Public Sub Case_MultiWindow()
    Dim w As Window
    Set w = ThisWorkbook.NewWindow
    Hold "multi"
    w.Close
End Sub

' @desc Changes the session window caption and holds Excel busy
Public Sub Case_Caption()
    ThisWorkbook.Windows(1).Caption = "Renamed"
    Hold "caption"
End Sub

' @desc Opens a second workbook and holds Excel busy for 40 seconds
Public Sub Case_SecondBookLong()
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Hold "second_long", 40
    wb.Close SaveChanges:=False
End Sub
