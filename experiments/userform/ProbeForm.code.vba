Option Explicit
' 한글 주석: 디자인과 별도로 에이전트가 작성하는 이벤트 코드
Private Sub cmdProbe_Click()
    Me.lblStatus.Caption = "안녕하세요, " & Me.txtName.Text
End Sub
