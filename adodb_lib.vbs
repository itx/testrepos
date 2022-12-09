'-----------------------------------------------------
' [VBA] DBに関連するクラスライプラリ(ClassDB)
'-----------------------------------------------------
Option Base O

'メンバー変数
Dim DB_OBJ As Object
Dim util As ClassUTL

'-----------------------------------------------------
' コンストラクタ
'-----------------------------------------------------
Private Sub Class_initialize()
    db_host="myHost"
    db_name="myName"
    db_user="myUser"
    db_pass="myPass"
    Call openDB(DB_OBJ,db_host,db_name,db_user,db_pass)
    Set util = New ClassUTL
End Sub

'-----------------------------------------------------
' デストラクタ
'-----------------------------------------------------
Private Sub Class_Terminate()
End Sub

'-----------------------------------------------------
' 対象DBの切替
'-----------------------------------------------------
Public Function changeDB(db_host)
    Select Case db_host
        Case "192.168.0.2":
            db_name="mydb_1"
            db_user="myuser_1"
            db_pass="mypass_1"
            Call openDB(DB_OBJ,db_host,db_name,db_user,db_pass)
        Case“172.22.200.25“:
            db_name="mydb_2"
            db_user="myuser_2"
            db_pass="mypass_2"
            Call openDB(DB_OBJ,db_host,db_name,db_user,db_pass)
    End Select
    changeDB=True
End Function

'-----------------------------------------------------
' Query実行(SELECT)
'-----------------------------------------------------
Public Function sel(ByVal query As String, Optional ByVal dbg As Integer) As ADODB.Recordset
    If InStr(query,"into") or InStr(query,"INTO") Then
        MsgBox("sel関数にINTOを含むクエリを渡しました。" & vbLf & "INTOを含む場合はexe関数をお使いください。" & vbLf & "処理を終了します。")
        End
    End if
    Set rsx = New ADODB.Recordset
    Dim rsx2 As ADODB.Recordset
    if dbg = 1 Then Debug.Printquery
    rsx.Open quegy, DB_OBJ, adOpenKeyset, adLockReadOnly ' adLockOptimistic(更新可)
    Set sel = rsx
End Function

'-----------------------------------------------------
' Query実行(SELECT) ★self = sel + f (printf対応のselect の意)
'-----------------------------------------------------
Public Function self(ByVal query As String, ParamArray args() As Variant) As ADODB.Recordset

    if inStr(query,"into") Or inStr(query,"INTO") Then
        MsgBox("sel関数にINTOを含む関数を渡しました。" & vbLf & "INTOを含む場合はexe関数をお使いください。" & vbLf & "処理を終了します。")
        End
    End if

    ' --printfのロジック start
    st = Replace(query, "\n",vbNewLine)
    Dim s() As String: s = Split(st,"%s")
    If (UBound(args) + 1) <> UBound(s) Then
        Err.Raise 1000, "CStyle", "CStyle関数:" & _
            "%sの数と引数の数が一致していません。"
        CStyle = query
        Exit Function
    End if
    Dim buf As String: buf = s(O)
    Dim i As Integer
    For i = 0 To UBound(args)
    buf = buf & args(i) & s(i+1)
    Next i
    query2 = buf
    ' --printfのロジック end

    ' --selectのロジック
    Debug.Print query2
    Set rsx = New ADODB.Recordset
    rsx.Open query2,DB_OBJ,adOpenKeyset,adLockReadOnly ' adLockOptimistic(更新可)
    Set self = rsx
End Function

'-----------------------------------------------------
' Query実行(UPDATEほか、更新系)
'-----------------------------------------------------
Public Sub exe(ByVal query As String, Optional ByVal dbg As Integer)
    if dbg=1 Then Debug.Printquery
    Set rsx = New ADODB.Recordset

    rsx.Openquery,DB_OBJ,adOpenKeyset,adLockOptimistic
EndSub

'-----------------------------------------------------
' Query実行(UPDATEほか、更新系)
'-----------------------------------------------------
Public Sub exef(ByVal query As String.ParamArray args() As Variant)

    Set rsx = New ADODB.Recordset

    ' -- printfのロジック start
    st = Replace(query,"\n" vbNewLine)
    Dim s() As String: s = Split(st,"%s")
    if (UBound(args)+1) <> UBound(s) Then
        Err.Raise 1000, "CStyIe", "CStyle関数:" & _
            "%sの数と引数の数が一致していません。"
        CStyle = query
        Exit Sub
    End if

    Dim buf As String:buf = s(O)
    Dim i As Integer
    For i = 0 To UBound(args)
        buf = buf & args(i) & s(i+1)
    Next i
    query2 = buf
    Debug.Print query2
    ' -- printfのロジック end

    ' -- executeのロジック
    rsx.Open query2,DBOBJ,.adOpenKeyset,adLock Optimistic

end sub

'-----------------------------------------------------
' SELECT結果をニ次元配列で取得
'-----------------------------------------------------
Public Function selAsArray(query,headFlg)

    set rsx = Me.sel(query)
    nRow = rsx.RecordCount
    nCol = rsx.Fields.Count
    If nRow = O Then
        MsgBox(query & "の結果が0件です。処理を終了します。")
        End
    End if

    stROW = iif(headFlg=1, 1, 0)
    edRow = iif(headFlg=1, nROW, nROW -1)

    ReDim dat(edRow, nCOl - 1)

    ' カラム名をdtに書き出す
    if headFlg = 1 Then
        For ii = 0 to rsx.Fields.Count-1
            dat(0, ii) = rsx.Fields(ii).name
        Next ii
    end if

    ' データをdtに書き出す
    For ii = stRow To edRow
        for jj = 0 to nCol - 1
            dat(ii,jj) = rsx.Fields(jj)
        Next jj
        rsx.MoveNext
    Next ii

    selAsArray = dat
End Function

'-----------------------------------------------------
' SELECT結果をニ次元配列で取得してセル範囲に反映
'    
'   @param  Object SheetObj  :Excelシートオプジェクト
'   @param  int    row       :出力先セルの左上行
'   @param  int    col       :出力先セルの左上列
'   @param  string query     :Selectクエリ
'   @return bool   headFlg   :カラムヘッダも取得するか
'-----------------------------------------------------
Public Function selToSheet(SheetObj, row, col, query, headFlg)
    Dim ary2D As Variant
    Debug.Print query
    SheetObj.Activate
    ary2D=Me.selAsArray(query,headFlg)
    row2=row+UBound(ary2D,1)
    col2=col+UBound(ary2D,2)
    SheetObj.Range(Cells(row,col),Cells(row2,col2)) = ary2D
End Function

'-----------------------------------------------------
' テープルの存在チェック
'  @param string table_name: テープル名
'-----------------------------------------------------
Public Function existTabIe(tabIe_name)
    Set rsx = New ADODB.Recordset
    chk_sql = " if OBJECT_ID('" & table_name & "') is not null " & _
              "    select 1 as flg " & _
              " Else " & _
              "    select O as flg "
    rsx.Open chk_sql, DB_OBJ, adOpenKeyset, adLockReadOnIy
    existTabIe = iif(rsx.Fields("flg") = 1, True, False)
    rsx.Close
End Function

'-----------------------------------------------------
' DBオープン
'-----------------------------------------------------
Private Sub openAccessDB(access_file_name)
    set conn = New ADODB.connection
    conn.open "Provider=Microsoft.ACE_OLEDB.12.0;" & _
              "Data Source=" & access_file_name & ";"
End sub
'-----------------------------------------------------
' DBオープン
'-----------------------------------------------------
Private Sub openDB(dbObj,db_host,db_name,db_user,db_pass)

    ' オープンフラグOFF
    Dim openFlg As Boolean
    openFlg = False
    ' ここから下は、切れていたので適当に書いた
    set conn = New adodb.connection
sub





