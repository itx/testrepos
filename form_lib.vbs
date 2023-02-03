'---------------------------------------------------------------
'
' フォームを親ウィンドウの中央に表示（マルチディスプレイ対応）
'
'    このメソッドをフォームのUserForm_initialized() で呼べばOK
'
'---------------------------------------------------------------
Private Sub UFPositionCenter()
    Me.StartUpPosition = 0    ' 0=指定なし  1=Form が属する項目の中央,  2=画面全体の中央,  3=画面の左上隅
    Me.Top = ActiveWindow.Top + ((ActiveWindow.Height - Me.Height) / 2)
    Me.Left = ActiveWindow.Left + ((ActiveWindow.Width - Me.Width) / 2)
End Sub
