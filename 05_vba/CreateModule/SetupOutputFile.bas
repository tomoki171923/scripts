Attribute VB_Name = "SetupOutputfile"
Public strFileName As String
Public WriteStream As Object

'------------------------------'
'テキストストリームのオブジェクト設定（Open）
'------------------------------'
Public Function OpenTextStream()
    ' テキストストリームのオブジェクト作成 & 設定(エンコード：UTF8、改行コード：LF)
    Set WriteStream = CreateObject("ADODB.Stream")
    With WriteStream
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
    End With

    ' テキストストリームのオブジェクトを開く
    WriteStream.Open
End Function

'------------------------------'
'テキストストリームのオブジェクト設定（Close）
'------------------------------'
Public Function CloseTextStream()
    ' テキストストリームのオブジェクトを閉じる
    WriteStream.Close
    ' オブジェクトクリア
    Set WriteStream = Nothing
End Function

'------------------------------'
'ファイル書き出し
'------------------------------'
Public Function FileOutput(ByVal strFileName As String)
    Dim objFSO As Object
    Dim Result As Long
    Dim strFullPath As String
    
    '引数で受け取ったファイル名でテキストストリーム書き出し
    'PATH指定
    strFullPath = ThisWorkbook.Worksheets("メイン").Range("C5").Value + strFileName
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' ファイル存在確認
    If objFSO.FileExists(strFullPath) = True Then
        Result = MsgBox("指定したフォルダに同名の" + strFileName + "が存在します。上書きしますか。", vbYesNo)
        If Result = 6 Then
            ' コンフィグファイル書き出し
            WriteStream.SaveToFile (strFullPath), adSaveCreateOverWrite
            MsgBox strFileName + "を上書きしました。"
        Else
            MsgBox strFileName + "の上書きを中止しました。"
            Set objFSO = Nothing
        End If
    Else
        ' コンフィグファイル書き出し
        WriteStream.SaveToFile (strFullPath), adSaveCreateOverWrite
        MsgBox strFileName + "を作成しました。"
    End If
    
    ' オブジェクトクリア
    Set objFSO = Nothing

End Function

