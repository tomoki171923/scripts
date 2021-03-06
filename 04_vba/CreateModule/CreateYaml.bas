Attribute VB_Name = "CreateYaml"
Option Explicit

Dim strPatternType As String
Dim strTargetSheet As String
Dim strEvFlag As String
Dim Count As Long
Dim ActiveFileColumn As Long
Dim ActiveRow As Long
Dim ActiveColumn As Long
Dim EndRow As Long
Dim strFileName As String
Dim strYamlCode As String
Dim strYamlKey As String
Dim strYamlValue As String
Dim strYamlName As String
Dim strYamlPattern As String
Dim TargetSheet As Worksheet
Dim WS As Worksheet
Dim FlagWS As Worksheet
Dim ListWS As Worksheet
Dim Rng As Object

'------------------------------'
'コントローラ
'------------------------------'
Sub ControllerYaml()
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
    Set ListWS = ThisWorkbook.Worksheets("コード一覧")

    'TargetListファイル作成
    Call CreateTargetList

    '開始位置設定
    StartColumn = 3
    StartRow = 5
    
    '各yamlファイル作成
    Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
        strFileName = ListWS.Cells(StartRow, StartColumn).Value
        Call JudgePattern(strFileName)
            
        StartColumn = StartColumn + 1
        Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
            strFileName = ListWS.Cells(StartRow, StartColumn).Value
            Call JudgePattern(strFileName)
            StartColumn = StartColumn + 1
        Loop
        '初期化
        StartColumn = 3
        strFileName = ""
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'パターン判定
'------------------------------'
Private Function JudgePattern(ByVal strFileName As String)
    Dim i As Long
    Dim j As Long
    Dim PatternCell As Object
    
    'ワークシ ートオブジェクト設定＆クリア
    Set FlagWS = ThisWorkbook.Worksheets("処理シート")
    FlagWS.Range("B4:D999999").ClearContents

    i = 4
    '処理シートに書き出し
    For Each TargetSheet In ThisWorkbook.Worksheets
        'パターン名検索
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not PatternCell Is Nothing Then
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
            
            '対象ファイル列に文字列が含んでいるか判定
            If JudgeSheet(strFileName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = strPatternType
                i = i + 1
            End If
        End If
    Next

    '対象ファイルの列に設定値が含まれていなかった場合、処理を終了
    If FlagWS.Range("B4:B999999").Find(what:="*") Is Nothing Then
        Exit Function
    End If

    'テキストストリームのオブジェクトを設定する（Open）
    Call OpenTextStream

    'ヘッダー作成
    Call CreateHeader(strFileName)
    
    j = 4
    Do While FlagWS.Cells(j, 2) <> ""
        strTargetSheet = FlagWS.Cells(j, 2)
        strPatternType = FlagWS.Cells(j, 3)
        strYamlKey = SetYamlKey(strTargetSheet, strPatternType)
       
        'ファイル名設定
        Select Case strPatternType
            Case "A"
               Call CreatePatternA(strTargetSheet, strFileName, strYamlKey)
            Case "B"
               Call CreatePatternB(strTargetSheet, strFileName, strYamlKey)
            Case "C"
               Call CreatePatternC(strTargetSheet, strFileName, strYamlKey)
            Case "D"
               Call CreatePatternD(strTargetSheet, strFileName, strYamlKey)
            Case "E"
               Call CreatePatternE(strTargetSheet, strFileName, strYamlKey)
            Case "F"
               Call CreatePatternF(strTargetSheet, strFileName, strYamlKey)
            Case "G"
               Call CreatePatternG(strTargetSheet, strFileName, strYamlKey)
            Case "H"
               Call CreatePatternH(strTargetSheet, strFileName, strYamlKey)
            Case "I"
                Call CreatePatternI(strTargetSheet, strFileName, strYamlKey)
            Case "J"
                Call CreatePatternJ(strTargetSheet, strFileName, strYamlKey)
           Case Else
        End Select
       j = j + 1
    Loop
            
    '書き出し関数呼び出し
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
    
    'テキストストリームのオブジェクトを設定する（Close）
    Call CloseTextStream
    
End Function


'------------------------------'
'処理シート判定
'------------------------------'
Public Function JudgeSheet(strFileName As String) As Boolean
    Dim strSerchTarget As String

    '対象ファイルの列を設定
    ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = 99999
        
    '対象ファイルの列に設定値が含めれているか判定
    strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
    Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
    If Rng Is Nothing Then
        JudgeSheet = False
    Else
        JudgeSheet = True
    End If

End Function


'------------------------------'
'エスケープ文字付加
'------------------------------'
Private Function SetEscape(ByVal strTargetVal As String, ByVal strTargetPat As String) As String
    '変数宣言
    Dim EscapeList As Variant
    Dim el As Variant
    
    Select Case strTargetPat
        Case "grep"
            'エスケープ対象の文字列を格納（シェル）
            EscapeList = Array("[", "]")
        Case "match"
            'エスケープ対象の文字列を格納（Ruby）
            EscapeList = Array("$", "*", """")
        Case "block"
            'エスケープ対象の文字列を格納（ヒアドキュメント）
            EscapeList = Array("""")
        Case Else
    End Select
    
    'エスケープ文字を付加
    For Each el In EscapeList
        strTargetVal = Replace(strTargetVal, el, "�" + el)
    Next
    
    SetEscape = strTargetVal
End Function


'------------------------------'
'セル内改行コード変換
'------------------------------'
Private Function SetLineBreak(ByVal strTarget As String) As String
      SetLineBreak = Replace(strTarget, vbLf, "�n")
End Function

               
'------------------------------'
'Yaml_Key設定
'------------------------------'
Private Function SetYamlKey(ByVal strTargetSheet As String, ByVal strPatternType As String) As String
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)

    Select Case strPatternType
        Case "A", "B", "C", "D", "H", "J"
            strYamlKey = WS.Cells(6, 3).Value
            strYamlKey = Mid(strYamlKey, InStrRev(strYamlKey, "/") + 1)
            strYamlKey = Replace(strYamlKey, ".", "_")
            strYamlKey = Replace(strYamlKey, "-", "_")
            strYamlKey = StrConv(strYamlKey, vbNarrow + vbProperCase)

        Case "E"
            If strTargetSheet = "【グループ】" Then
                strYamlKey = "Group"
            ElseIf strTargetSheet = "【ユーザ】" Then
                strYamlKey = "User"
            ElseIf strTargetSheet = "【ファイル配布】" Then
                strYamlKey = "File"
            ElseIf strTargetSheet = "【ディレクトリ作成】" Then
                strYamlKey = "Directory"
            ElseIf strTargetSheet = "【パッケージ】" Then
                strYamlKey = "Rpm"
            End If
            
        Case "F"
            strYamlKey = "BootService"
            
        Case "G"
            strYamlKey = "XinetdService"
            
        Case "I"
            strYamlKey = "EnvironmentVariable"
            
        Case Else
    End Select
    SetYamlKey = strYamlKey
End Function

'------------------------------'
'TargetListファイル作成
'------------------------------'
Private Function CreateTargetList()
    Dim i As Long
    
    'テキストストリームのオブジェクトを設定する（Open）
    Call OpenTextStream
    
    'コード書き出し
    strYamlCode = "#############################################"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Name                : targetList.yml"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Overview            : This is Target Server List of spec files"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Creation date       : " + Format(Date, "yyyy.mm.dd")
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Correction history  :"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Constraint          :"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "#############################################"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "node:"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'ワークシートオブジェクト設定
    Set ListWS = ThisWorkbook.Worksheets("コード一覧")
    
    '開始位置設定
    ActiveColumn = 4
    ActiveRow = 5
    
    Do While ListWS.Cells(ActiveRow, ActiveColumn).Value <> ""
        Do While ListWS.Cells(ActiveRow, ActiveColumn).Value <> ""
            strYamlCode = "  - name: '" + ListWS.Cells(ActiveRow, ActiveColumn).Value + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            strYamlCode = "    role: '" + ListWS.Cells(ActiveRow, 3).Value + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            strYamlCode = "    spec:"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ワークシートオブジェクト設定
            Set WS = ThisWorkbook.Worksheets("レシピ一覧")
            
            '開始行設定
            i = 4
    
            '終了行設定
            EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
            
            'CookBook名取得
            Do While i < EndRow
                If WS.Cells(i, 3).Value <> "" Then
                    strYamlCode = "     - '" + WS.Cells(i, 3).Value + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                End If
                i = i + 1
            Loop
            ActiveColumn = ActiveColumn + 1
        Loop
        '初期化
        ActiveColumn = 4
        
        ActiveRow = ActiveRow + 1
    Loop
    
    '書き出し関数呼び出し
    Call FileOutput("targetList.yml")
    
    'テキストストリームのオブジェクトを設定する（Close）
    Call CloseTextStream

End Function

'------------------------------'
'ヘッダー作成
'------------------------------'
Private Function CreateHeader(ByVal strFileName As String)
    'コード書き出し
    strYamlCode = "#############################################"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Name                : " + strFileName + ".yml"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Overview            : This is Variable list of spec files"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Creation date       : " + Format(Date, "yyyy.mm.dd")
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Correction history  :"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "# Constraint          :"
    WriteStream.WriteText strYamlCode, adWriteLine
    strYamlCode = "#############################################"
    WriteStream.WriteText strYamlCode, adWriteLine
End Function

'------------------------------'
'パターンAコード生成
'------------------------------'
Private Function CreatePatternA(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml「name」key名の列を取得
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    Do While WS.Cells(ActiveRow, 12) <> "以上"
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'コメントアウト文判定
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
                
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml「name」取得
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'コード書き出し（name）
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'パターンBコード生成
'------------------------------'
Private Function CreatePatternB(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ブロック文、コメントアウト文の判定
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'コード書き出し（lines）
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（grep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'コード書き出し（grep_val）
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'パターンCコード生成
'------------------------------'
Private Function CreatePatternC(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml「name」key名の列を取得
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '変更項目処理
    '------------------------------'
    '変更項目処理開始行設定
    ActiveRow = 6
    
    '変更項目処理終了行設定
    EndRow = WS.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'コメントアウト文判定
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml「name」取得
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'コード書き出し（name）
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '追加項目処理
    '------------------------------'
    '追加項目処理開始行設定
    ActiveRow = ActiveRow + 3
    
    '追加項目処理終了行設定
    EndRow = WS.Cells.Find("以上（追加項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ブロック文、コメントアウト文の判定
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（name）
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'コード書き出し（lines）
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（grep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'コード書き出し（grep_val）
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'パターンDコード生成（grub.conf）
'------------------------------'
Private Function CreatePatternD(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '変数宣言
    Dim lines() As String
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml「name」key名の列を取得
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '変更項目処理
    '------------------------------'
    '変更項目処理開始行設定
    ActiveRow = 6
    
    '変更項目処理終了行設定
    EndRow = WS.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'コメントアウト文判定
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml「name」取得
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'コード書き出し（name）
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '追加項目(kernel行)
    '------------------------------'
    '追加項目処理開始行設定
    ActiveRow = ActiveRow + 3
    
    '追加項目処理終了行設定
    EndRow = WS.Cells.Find("以上（追加項目(kernel行)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コメントアウト文の判定
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（name）
            strYamlCode = "    :name: 'kernel'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（grep_val）
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '追加項目(ファイル末尾)
    '------------------------------'
    '追加項目処理開始行設定
    ActiveRow = ActiveRow + 3
    
    '追加項目処理終了行設定
    EndRow = WS.Cells.Find("以上（追加項目(ファイル末尾)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ブロック文、コメントアウト文の判定
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（name）
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'コード書き出し（lines）
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（grep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'コード書き出し（grep_val）
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'パターンEコード生成
'------------------------------'
Private Function CreatePatternE(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    Dim HashArray() As String
    Dim HashCount As Long
    Dim HashColumn As Long
    Dim HashRow As Long
    Dim strHashName As String
    Dim strHashValue As String
    Dim ha As Variant
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Hash個数取得
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash名位置取得
    HashRow = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine

    Select Case strTargetSheet
        Case "【ユーザ】"
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
                    'コード書き出し（-key）
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlパターン判定
                    'コメントアウト文か判定
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'コード書き出し（pattern）
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '初期化
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash数分行ループ処理
                    Do While Count < HashCount
                        'コード書き出し（Hash名）
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn + Count).Value
                        strHashValue = SetEscape(strHashValue, "match")
                    
                        strHashName = WS.Cells(HashRow, HashColumn + Count).Value
                        If strHashName = "secondary_group_name" Then
                            strYamlCode = "    :secondary_group_name:"
                            WriteStream.WriteText strYamlCode, adWriteLine
                        
                            HashArray = Split(strHashValue, ",")
                            For Each ha In HashArray
                                strYamlCode = "      - '" + ha + "'"
                                WriteStream.WriteText strYamlCode, adWriteLine
                            Next
                        ElseIf strHashName = "password" Then
                            GoTo Continue
                        Else
                            strYamlCode = "    :" + WS.Cells(HashRow, HashColumn + Count).Value + ": '" + strHashValue + "'"
                            WriteStream.WriteText strYamlCode, adWriteLine
                        End If
Continue:
                        Count = Count + 1
                    Loop
                End If
                ActiveRow = ActiveRow + 1
            Loop
        Case "【パッケージ】"
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
                    'コード書き出し（-key）
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlパターン判定
                    'コメントアウト文か判定
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'コード書き出し（pattern）
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '初期化
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash数分行ループ処理
                    Do While Count < HashCount
                        'コード書き出し（Hash名）
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn + Count).Value
                        strHashValue = SetEscape(strHashValue, "match")
                    
                        strHashName = WS.Cells(HashRow, HashColumn + Count).Value
                        If strHashName = "rpm_name" Then
                            strYamlCode = "    :" + WS.Cells(HashRow, HashColumn + Count).Value + ": '" + strHashValue + "'"
                            WriteStream.WriteText strYamlCode, adWriteLine
                        End If
                        Count = Count + 1
                    Loop
                End If
                ActiveRow = ActiveRow + 1
            Loop
        
        Case Else
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
                    'コード書き出し（-key）
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlパターン判定
                    'コメントアウト文か判定
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'コード書き出し（pattern）
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '初期化
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash数分行ループ処理
                    Do While Count < HashCount
                        'コード書き出し（Hash名）
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn + Count).Value
                        strHashValue = SetEscape(strHashValue, "match")
                        strYamlCode = "    :" + WS.Cells(HashRow, HashColumn + Count).Value + ": '" + strHashValue + "'"
                        WriteStream.WriteText strYamlCode, adWriteLine
                        Count = Count + 1
                    Loop
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
    End Select
End Function

'------------------------------'
'パターンFコード生成（自動起動サービス）
'------------------------------'
Private Function CreatePatternF(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "以上"
        If WS.Cells(ActiveRow, 3) <> "" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（name）
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            If WS.Cells(ActiveRow, ActiveFileColumn) = "" Or WS.Cells(ActiveRow, ActiveFileColumn) = "-" Then
                'コード書き出し（pattern）
                strYamlCode = "    :pattern: 'noexist'"
                WriteStream.WriteText strYamlCode, adWriteLine
        
            ElseIf WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                'コード書き出し（pattern）
                strYamlCode = "    :pattern: 'exist'"
                WriteStream.WriteText strYamlCode, adWriteLine
            
                Count = 0
                Do While Count < 7
                    'コード書き出し（runlevel）
                    strYamlCode = "    :runlevel" + CStr(Count) + ": '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, "match") + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                    Count = Count + 1
                Loop
           
            End If
        End If
        ActiveRow = ActiveRow + 1
    Loop
End Function

'------------------------------'
'パターンGコード生成（xinetdサービス）
'------------------------------'
Private Function CreatePatternG(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "以上"
        If WS.Cells(ActiveRow, 3) <> "" Then
    
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'コード書き出し（name）
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'コード書き出し（pattern）
            Select Case WS.Cells(ActiveRow, ActiveFileColumn).Value
                Case "on"
                    strYamlCode = "    :pattern: 'on'"
                Case "off"
                    strYamlCode = "    :pattern: 'off'"
                Case "", "-"
                    strYamlCode = "    :pattern: 'noexist'"
                Case Else
            End Select
            WriteStream.WriteText strYamlCode, adWriteLine
        End If
        
        ActiveRow = ActiveRow + 1
    Loop
End Function

'------------------------------'
'パターンHコード生成（Hosts）
'------------------------------'
Private Function CreatePatternH(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 4) <> "以上"
        If WS.Cells(ActiveRow, 4) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            ActiveColumn = 5
            Do While ActiveColumn < 8
                If WS.Cells(ActiveRow, ActiveColumn) <> "" Then
                    'コード書き出し（-key）
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
            
                    'コード書き出し（ipaddress）
                    strYamlCode = "    :ipaddress: '" + WS.Cells(ActiveRow, 4) + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'コード書き出し（Hostname)
                    strYamlCode = "    :hostname: '" + WS.Cells(ActiveRow, ActiveColumn) + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                End If
                ActiveColumn = ActiveColumn + 1
            Loop
        End If
        ActiveRow = ActiveRow + 1
    Loop
End Function

'------------------------------'
'パターンIコード生成（ユーザ環境変数）
'------------------------------'
Private Function CreatePatternI(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'コード書き出し（key）
    If strEvFlag <> strFileName Then
        strYamlCode = strYamlKey + ":"
        WriteStream.WriteText strYamlCode, adWriteLine
        strEvFlag = strFileName
    End If
    
    'コード書き出し（-key）
    strYamlCode = "  - " + LCase(strYamlKey) + " :"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'コード書き出し（path）
    strYamlCode = "    :path: '" + WS.Cells(6, 5) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'コード書き出し（file_name）
    strYamlCode = "    :file_name: '" + WS.Cells(6, 6) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine

    'コード書き出し（value）
    strYamlCode = "    :value:"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 12) <> "以上"
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "○" Then
            'コード書き出し（-val）
            strYamlCode = "      - val :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            'ブロック文、コメントアウト文の判定
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "        :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
                
            If strYamlPattern = "block" Then
                'コード書き出し（lines）
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "        :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（grep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "        :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'コード書き出し（match_val）
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "        :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'コード書き出し（grep_val）
                strYamlCode = "        :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine

                'コード書き出し（match_val）
                strYamlCode = "        :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If

        End If
        
        ActiveRow = ActiveRow + 1
    Loop
End Function


'------------------------------'
'パターンJコード生成（sysctl.conf）
'------------------------------'
Private Function CreatePatternJ(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'コード書き出し（key）
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml「name」key名の列を取得
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '変更項目処理
    '------------------------------'
    '変更項目処理開始行設定
    ActiveRow = 6
    
    '変更項目処理終了行設定
    EndRow = WS.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'コメントアウト文判定
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml「name」取得
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'コード書き出し（name）
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '追加項目処理
    '------------------------------'
    '追加項目処理開始行設定
    ActiveRow = ActiveRow + 3
    
    '追加項目処理終了行設定
    EndRow = WS.Cells.Find("以上（追加項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
            'コード書き出し（-key）
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コメントアウト文の判定
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'コード書き出し（pattern）
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（name）
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（grep_val）
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'コード書き出し（match_val）
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function


Sub test()
    strTargetSheet = "【起動スクリプト】rc.local"
    strFileName = "stprdb01"
    strYamlKey = "Rc_local"
    'フォルダ判定
    OutputPath = ThisWorkbook.Worksheets("メイン").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "出力先フォルダに誤りがあります。", vbCritical
        Exit Sub
    End If

    'テキストストリームのオブジェクトを設定する（Open）
    Call OpenTextStream
    Call CreatePatternC(strTargetSheet, strFileName, strYamlKey)
        
    '書き出し関数呼び出し
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
        
    'テキストストリームのオブジェクトを設定する（Close）
    Call CloseTextStream

End Sub

