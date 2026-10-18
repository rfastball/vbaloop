Attribute VB_Name = "modMemoDemo"
Option Explicit

Public Sub ShowDemo()
    frmMemo.Show vbModeless
End Sub

Public Sub Test_MemoDemo()
    Load frmMemo
    frmMemo.Show vbModeless
    frmMemo.cmdPreview.Value = True
    LogShim.AssertEq "memo preview", "김하늘님의 메모" & vbCrLf & "JSON과 VBA로 첫 폼을 만들었습니다.", frmMemo.lblResult.Caption
    frmMemo.cmdReset.Value = True
    LogShim.AssertEq "reset name", "", frmMemo.txtName.Text
    LogShim.AssertEq "reset memo", "", frmMemo.txtMemo.Text
    frmMemo.cmdPreview.Value = True
    LogShim.AssertEq "empty name", "이름을 먼저 입력해 주세요.", frmMemo.lblResult.Caption
    frmMemo.txtName.Text = "김하늘"
    frmMemo.cmdPreview.Value = True
    LogShim.AssertEq "empty memo", "미리볼 메모를 입력해 주세요.", frmMemo.lblResult.Caption
    Unload frmMemo
End Sub
