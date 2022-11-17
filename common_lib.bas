'-----------------------------------------------------
' ファイルの存在チェック(フォルダには使えない)
'-----------------------------------------------------
Public Function isExist(filePath As String)
    If Dir(filePath) <> "" Then
        isExist = True
    Else
        isExist = False
    End If
End Function

'-----------------------------------------------------
' フォルダがなければ作成
'-----------------------------------------------------
Public Function makeDirectory(path As String)
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
End Function

'-----------------------------------------------------
' Dir関数の結果を配列で返す
'  usage : myfiles = getFileArray("c:\temp\*.sys")
'-----------------------------------------------------
Public Function getFileArray(ByVal targetPath As String, Optional ByVal attr As VbFileAttribute = VbFileAttribute.vbNormal) As String()
    Dim resultArray() As String
    Dim result As String
    Dim suffix As Long
    suffix = 0
    
    result = Dir(targetPath, attr)
    ReDim resultArray(suffix)
    resultArray(suffix) = result
    
    Do While result <> ""
        result = Dir()
        If result <> "" Then
            suffix = suffix + 1
            ReDim Preserve resultArray(suffix)
            resultArray(suffix) = result
        End If
    Loop
    getFileArray = resultArray
End Function

