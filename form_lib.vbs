'---------------------------------------------------------------
'
' フォームを親ウィンドウの中央に表示（マルチディスプレイ対応）
'
'    このメソッドをフォームのUserForm_initialized() で呼べばOK
'
'---------------------------------------------------------------
Private Sub UFPositionCenter()
  '**ユーザーフォームを親ウィンドウの中央に表示する
  
  '**変数(T=Top,L=Left,W=Width,H=Height,AW=ActiveWindow,UF=UserForm)
  Dim T_AW As Long, L_AW As Long, W_AW As Long, H_AW As Long
  Dim T_UF As Long, L_UF As Long, W_UF As Long, H_UF As Long
  
  '**親ウィンドウの位置とサイズを取得
  With ActiveWindow
    T_AW = .Top
    L_AW = .Left
    W_AW = .Width
    H_AW = .Height
  End With
  
  '**UFのサイズを取得
  W_UF = Me.Width
  H_UF = Me.Height

  '**UFの表示位置を計算
  T_UF = T_AW + ((H_AW - H_UF) / 2)
  L_UF = L_AW + ((W_AW - W_UF) / 2)
  
  '**UFの表示位置を設定
  Me.StartUpPosition = 0    ' 0=指定なし  1=Form が属する項目の中央,  2=画面全体の中央,  3=画面の左上隅
      '**Top,Left指定時に必須(ないとLeftがずれる)
  Me.Top = T_UF
  Me.Left = L_UF
End Sub