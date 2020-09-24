Attribute VB_Name = "CreateErb"
Dim TargetSheet As Worksheet
Dim strPatternType As String


Sub ControllerErb()
    '変数宣言
    Dim StartColumn As Long
    Dim StartRow As Long
    
    'フォルダ判定
    OutputPath = ThisWorkbook.Worksheets("メイン").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "出力先フォルダに誤りがあります。", vbCritical
        Exit Sub
    End If
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets("コード一覧")

    '開始位置設定
    StartColumn = 3
    StartRow = 5
    
    Do While WS.Cells(StartRow, StartColumn).Value <> ""
        strTargetName = WS.Cells(StartRow, StartColumn).Value
        Call Tyuukann(strTargetName)
            
        StartColumn = StartColumn + 1
        Do While WS.Cells(StartRow, StartColumn).Value <> ""
            strTargetName = WS.Cells(StartRow, StartColumn).Value
            Call Tyuukann(strTargetName)
            StartColumn = StartColumn + 1
        Loop
        '初期化
        StartColumn = 3
        strTargetName = ""
        StartRow = StartRow + 1
    Loop
End Sub
    
Private Function Tyuukann(ByVal strTargetName As String)
    'ワークシ ートオブジェクト設定＆クリア
    Set FlagWS = ThisWorkbook.Worksheets("処理シート")
    FlagWS.Range("B4:D999999").ClearContents

    i = 4
    '処理シートに書き出し
    For Each TargetSheet In Worksheets
        If InStr(TargetSheet.Name, "hosts") > 0 Then
            If CheckSheet(strTargetName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = "hosts"
                i = i + 1
            End If
        ElseIf InStr(TargetSheet.Name, "ユーザ環境変数（個別）") > 0 Then
            If CheckSheet(strTargetName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = "ユーザ環境変数"
                i = i + 1
            End If
        End If
    Next

    '対象ファイルの列に設定値が含まれていなかった場合、処理を終了
    If FlagWS.Range("B4:B999999").Find(what:="*") Is Nothing Then
        Exit Function
    End If

    j = 4
    Do While FlagWS.Cells(j, 2) <> ""
        'テキストストリームのオブジェクトを設定する（Open）
        Call OpenTextStream

        strTargetSheet = FlagWS.Cells(j, 2)
        strPatternType = FlagWS.Cells(j, 3)
        'ファイル名設定
        Select Case strPatternType
            Case "hosts"
                Call CreateHosts(strTargetSheet, strTargetName)
            Case "ユーザ環境変数"
                Call CreateEnviromentVal(strTargetSheet, strTargetName)
            Case Else
        End Select
            
        ' 書き出し関数呼び出し
        strFileName = strFileName + ".erb"
        Call FileOutput(strFileName)
        
        'テキストストリームのオブジェクトを設定する（Close）
        Call CloseTextStream
        
        j = j + 1
    Loop
    
End Function


Function CreateHosts(ByVal strTargetSheet As String, ByVal strTargetName As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    'ヘッダー書き出し
    WriteStream.WriteText "127.0.0.1   localhost localhost.localdomain localhost4 localhost4.localdomain4", adWriteLine
    WriteStream.WriteText "#::1         localhost localhost.localdomain localhost6 localhost6.localdomain6", adWriteLine
    WriteStream.WriteText "", adWriteLine
    WriteStream.WriteText "", adWriteLine
            
    '追加項目列ループ処理
    Do While WS.Cells(ActiveRow, 4) <> "以上"
        If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
            strChefCode = WS.Cells(ActiveRow, 4).Value + vbTab + WS.Cells(ActiveRow, 5).Value + vbTab + WS.Cells(ActiveRow, 6).Value + vbTab + WS.Cells(ActiveRow, 7).Value + vbTab + WS.Cells(ActiveRow, 8).Value
            WriteStream.WriteText strChefCode, adWriteLine
  
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '出力ファイル名設定
    strFileName = strTargetName + "_hosts"
    
End Function




Function CreateEnviromentVal(ByVal strTargetSheet As String, ByVal strTargetName As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    'ヘッダー書き出し
    Select Case WS.Cells(6, 6)
        Case ".bash_profile"
            WriteStream.WriteText "# .bash_profile", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# Get the aliases and functions", adWriteLine
            WriteStream.WriteText "if [ -f ~/.bashrc ]; then", adWriteLine
            WriteStream.WriteText "# .bash_profile", adWriteLine
            WriteStream.WriteText "    . ~/.bashrc", adWriteLine
            WriteStream.WriteText "fi", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific environment and startup programs", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "PATH=$PATH:$HOME/bin", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "export PATH", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".bashrc"
            WriteStream.WriteText "# .bashrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# Source global definitions", adWriteLine
            WriteStream.WriteText "if [ -f /etc/bashrc ]; then", adWriteLine
            WriteStream.WriteText "    . /etc/bashrc", adWriteLine
            WriteStream.WriteText "fi", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".tcshrc"
            WriteStream.WriteText "# .tcshrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "set prompt='[%n@%m %c]# '", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".cshrc"
            WriteStream.WriteText "# .cshrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case Else
    End Select
    
    'root判定処理
    If WS.Cells(6, 3) = "root" Then
        WriteStream.WriteText "export PATH=chefdk/frhu", adWriteLine
        WriteStream.WriteText "", adWriteLine
    End If
    
    
    '追加項目列ループ処理
    Do While WS.Cells(ActiveRow, 12) <> "以上"
        If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
            strChefCode = WS.Cells(ActiveRow, 12).Value
            WriteStream.WriteText strChefCode, adWriteLine
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '出力ファイル名設定
    strFileName = strTargetName + "_" + WS.Cells(6, 3).Value + WS.Cells(6, 6).Value

End Function

'------------------------------'
'処理シート判定
'------------------------------'
Public Function CheckSheet(strTargetName As String) As Boolean

    '対象ファイルの列を設定
    ActiveFileColumn = TargetSheet.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = TargetSheet.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
        
    '対象ファイルの列に設定値が含めれているか判定
    strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
    Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="○")
    If Rng Is Nothing Then
        CheckSheet = False
    Else
        CheckSheet = True
    End If

End Function
