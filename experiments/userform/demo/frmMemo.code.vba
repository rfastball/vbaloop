Option Explicit

Private Sub UserForm_Initialize()
    Dim ctl As MSForms.Control
    Me.BackColor = RGB(246, 247, 244)
    For Each ctl In Me.Controls
        ctl.Font.Name = "맑은 고딕"
        ctl.Font.Size = 10
        ctl.ForeColor = RGB(35, 46, 44)
        If TypeName(ctl) = "Label" Then ctl.BackStyle = 0
    Next ctl
    Me.lblFootnote.Font.Size = 8
    Me.lblFootnote.ForeColor = RGB(80, 93, 88)
    Me.lblName.Font.Bold = True
    Me.lblMemo.Font.Bold = True
    Me.lblResult.ForeColor = RGB(49, 75, 61)
    Me.lblResult.WordWrap = True
    Me.txtName.SpecialEffect = fmSpecialEffectFlat
    Me.txtName.BorderStyle = fmBorderStyleSingle
    Me.txtName.BorderColor = RGB(178, 187, 179)
    Me.txtMemo.SpecialEffect = fmSpecialEffectFlat
    Me.txtMemo.BorderStyle = fmBorderStyleSingle
    Me.txtMemo.BorderColor = RGB(178, 187, 179)
    Me.txtMemo.ScrollBars = fmScrollBarsVertical
    Me.cmdReset.BackColor = RGB(234, 237, 231)
    Me.cmdClose.BackColor = RGB(234, 237, 231)
    Me.txtMemo.MultiLine = True
    Me.txtMemo.EnterKeyBehavior = True
    Me.cmdPreview.BackColor = RGB(24, 82, 65)
    Me.cmdPreview.ForeColor = RGB(255, 255, 255)
    Me.cmdPreview.Font.Bold = True
    Me.cmdPreview.Default = True
    Me.cmdPreview.Accelerator = "P"
    Me.cmdReset.Accelerator = "R"
    Me.cmdClose.Accelerator = "C"
    Me.cmdClose.Cancel = True
    Me.txtName.TabIndex = 0
    Me.txtMemo.TabIndex = 1
    Me.cmdPreview.TabIndex = 2
    Me.cmdReset.TabIndex = 3
    Me.cmdClose.TabIndex = 4
End Sub

Private Sub cmdPreview_Click()
    If Len(Trim$(Me.txtName.Text)) = 0 Then
        Me.lblResult.Caption = "이름을 먼저 입력해 주세요."
        Me.txtName.SetFocus
        Exit Sub
    End If
    If Len(Trim$(Me.txtMemo.Text)) = 0 Then
        Me.lblResult.Caption = "미리볼 메모를 입력해 주세요."
        Me.txtMemo.SetFocus
        Exit Sub
    End If
    Me.lblResult.Caption = Trim$(Me.txtName.Text) & "님의 메모" & vbCrLf & Trim$(Me.txtMemo.Text)
End Sub

Private Sub cmdReset_Click()
    Me.txtName.Text = ""
    Me.txtMemo.Text = ""
    Me.lblResult.Caption = "새 메모를 입력해 주세요."
    Me.txtName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
