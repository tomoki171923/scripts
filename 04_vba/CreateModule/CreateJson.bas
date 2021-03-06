Attribute VB_Name = "CreateJson"
'Option Explicit
    '変数宣言（グローバル）
    Dim TargetSheet As Worksheet
    Dim strTargetSheet As String
    Dim strFileType As String
    Dim strChefCode As String
    Dim strRecipeName As String
    Dim strCookbookName As String
    Dim strRoleName As String
    Dim strNodeName As String
    Dim ActiveFileColumn As Long
    Dim ActiveColumn As Long
    Dim ActiveRow As Long
    Dim EndRow As Long
    Dim strSerchTarget As String
    Dim AttributeRow As Long
    Dim AttributeColumn As Long
    Dim strAttributeName As String
    Dim strAttributeValue As String
    Dim strLastSheet As String
    Dim strPatternType As String
    Dim AttWrite As Long
    Dim HashWrite As Long
    Dim FlagWS As Worksheet


    Dim AttWriteI As Long
    
    '定数宣言（グローバル）
    Public Const Dquote As String = """"
    Public Const indent1 As String = "  "
    Public Const indent2 As String = "    "
    Public Const indent3 As String = "      "
    Public Const indent4 As String = "        "
    Public Const indent5 As String = "          "
    
'------------------------------'
'コントローラー
'------------------------------'
Sub ControllerJson()
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
    StartRow = 5
    StartColumn = 3
    
    Do While WS.Cells(StartRow, StartColumn).Value <> ""
        strRoleName = WS.Cells(StartRow, StartColumn).Value
        Call JudgePattern("Role", strRoleName, strNodeName)
        
        StartColumn = StartColumn + 1
        Do While WS.Cells(StartRow, StartColumn).Value <> ""
            strNodeName = WS.Cells(StartRow, StartColumn).Value
            Call JudgePattern("Node", strRoleName, strNodeName)
            StartColumn = StartColumn + 1
        Loop
        '初期化
        StartColumn = 3
        strNodeName = ""
        
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'パターン判定
'------------------------------'
Private Function JudgePattern(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
    'テキストストリームのオブジェクトを設定する（Open）
    Call OpenTextStream
    
    'ヘッダーを作成する
    Call CreateHeader(strFileType, strRoleName, strNodeName)
    
    'ファイル名設定
    Select Case strFileType
        Case "Role"
            strFileName = strRoleName
        Case "Node"
            strFileName = strNodeName
        Case Else
    End Select
       
    'ワークシ ートオブジェクト設定＆クリア
    Set FlagWS = ThisWorkbook.Worksheets("処理シート")
    FlagWS.Range("B4:D999999").ClearContents
       
    '処理シート設定
    Call JudgeSheets(strFileName)

    '最終処理シート名取得
    strLastSheet = FlagWS.Range("B3").End(xlDown).Value
        
    a = 4
    Do While FlagWS.Cells(a, 2) <> ""
        strTargetSheet = FlagWS.Cells(a, 2)
        strPatternType = FlagWS.Cells(a, 3)
    
        'パターン判定
        Select Case strPatternType
            Case "A"
                Call CreatePatternA(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "B"
                Call CreatePatternB(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "C"
                Call CreatePatternC(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "D"
                Call CreatePatternD(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "E"
                Call CreatePatternE(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "F"
                Call CreatePatternF(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "G"
                Call CreatePatternG(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "H"
                Call CreatePatternH(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case "I"
                Call CreatePatternI(strFileType, strTargetSheet, strFileName, strLastSheet)
            Case Else
        End Select
        a = a + 1
    Loop
    
    'フッターを作成する
    Call CreateFooter(strFileType)

    ' 書き出し関数呼び出し
    strFileName = strFileName + ".json"
    Call FileOutput(strFileName)

    'テキストストリームのオブジェクトを設定する（Close）
    Call CloseTextStream

End Function

'------------------------------'
'処理シート判定
'------------------------------'
Private Function JudgeSheets(ByVal strFileName As String) As String
    '変数宣言
    Dim PatternCell As Range
    i = 4
    'シートループ
    For Each TargetSheet In ThisWorkbook.Worksheets
        'パターン名検索
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        If PatternCell Is Nothing Then
            strPatternType = ""
        Else
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
        End If
    
        '開始行設定
        ActiveRow = 6
 
        '対象ファイルの列に設定値が含めれているか判定
        Select Case strPatternType
            Case "A"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
                
                '
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Not Rng Is Nothing Then
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If
        
            Case "B", "E", "H", "I"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
        
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="○")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If
                
            Case "C"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上（追加項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '対象ファイルの列に"○"が含まれているか判定
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="○")

                If Rng Is Nothing And strFlag = "off" Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If
                
            Case "D"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上（追加項目(kernel行)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '対象ファイルの列に"○"が含まれているか判定
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="○")

                If Rng Is Nothing And strFlag = "off" Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上（追加項目(ファイル末尾)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '対象ファイルの列に"○"が含まれているか判定
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="○")

                If Rng Is Nothing And strFlag = "off" Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If

            Case "F"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row

                'Templateの列を設定
                TemplateColumn = 7
    
                '対象ファイルの列に設定値が含まれているか判定
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn + 6).Address
                Set RngOn = TargetSheet.Range(strSerchTarget).Find(what:="?:on")
                Set RngOff = TargetSheet.Range(strSerchTarget).Find(what:="?:off")
                Set RngHyphen = TargetSheet.Range(strSerchTarget).Find(what:="-")
                If RngOn Is Nothing And RngOff Is Nothing And RngHyphen Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If
                
            Case "G"
                '対象ファイルの列を設定
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '終了行設定
                EndRow = TargetSheet.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
    
                '対象ファイルの列に設定値が含まれているか判定
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set RngOn = TargetSheet.Range(strSerchTarget).Find(what:="on")
                Set RngOff = TargetSheet.Range(strSerchTarget).Find(what:="off")
                If RngOn Is Nothing And RngOff Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                    FlagWS.Cells(i, 2).Value = TargetSheet.Name
                    FlagWS.Cells(i, 3).Value = strPatternType
                    FlagWS.Cells(i, 4).Value = strFlag
                    i = i + 1
                End If
               
            Case Else
        End Select
    Next
End Function

'------------------------------'
'エスケープ文字付加
'------------------------------'
Private Function AddEscape(ByVal strTarget As String) As String
    'ダブルクォーテーションがある場合エスケープ文字追加
    AddEscape = Replace(strTarget, """", "�""")
End Function


'------------------------------'
'ヘッダー作成
'------------------------------'
Private Function CreateHeader(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
        'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets("レシピ一覧")
    
    '開始行設定
    ActiveRow = 4
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            
            WriteStream.WriteText "{", adWriteLine
            strChefCode = indent1 + Dquote + "name" + Dquote + ": " + Dquote + strRoleName + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + Dquote + "json_class" + Dquote + ": " + Dquote + "Chef::Role" + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + Dquote + "chef_type" + Dquote + ": " + Dquote + "role" + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + Dquote + "run_list" + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
        
        
            'CookBook名取得
            Do While ActiveRow < EndRow
                If WS.Cells(ActiveRow, 3).Value <> "" Then
                    If WS.Cells(ActiveRow + 1, 3).Value <> "" Then
                        '処理行数設定
                        Line = 1
                    Else
                        '処理行数設定
                        LineCount = WS.Cells(ActiveRow, 3).End(xlDown).Row
                        Line = LineCount - ActiveRow
                    End If
                    strCookbookName = WS.Cells(ActiveRow, 3).Value
                
                    j = 1
                    'Recipe名取得
                    Do While j <= Line
                        strRecipeName = WS.Cells(ActiveRow, 4).Value
                        
                        If strRecipeName <> "" Then
                            'カンマ判定処理（run_list内）
                            If ActiveRow + 1 < EndRow Then
                                strChefCode = indent2 + Dquote + "recipe[" + strCookbookName + "::" + strRecipeName + "]" + Dquote + ","
                                WriteStream.WriteText strChefCode, adWriteLine
                            Else
                                strChefCode = indent2 + Dquote + "recipe[" + strCookbookName + "::" + strRecipeName + "]" + Dquote
                                WriteStream.WriteText strChefCode, adWriteLine
                            End If
                        End If
                        ActiveRow = ActiveRow + 1
                        j = j + 1
                    Loop
                End If
            Loop
            
            strChefCode = indent1 + "],"
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + Dquote + "default_attributes" + Dquote + ": {"
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent2 + Dquote + strCookbookName + Dquote + ": {"
            WriteStream.WriteText strChefCode, adWriteLine
        
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
        
            WriteStream.WriteText "{", adWriteLine
            strChefCode = indent1 + Dquote + "name" + Dquote + ": " + Dquote + strNodeName + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + Dquote + "run_list" + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent2 + Dquote + "role[" + strRoleName + "]" + Dquote
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent1 + "],"
            WriteStream.WriteText strChefCode, adWriteLine
            
            'CookBook名取得
            Do While WS.Cells(ActiveRow, 3).Value <> "以上"
                strCookbookName = WS.Cells(ActiveRow, 3).Value
                
                If strCookbookName <> "" Then
                    strChefCode = indent1 + Dquote + strCookbookName + Dquote + ": {"
                    WriteStream.WriteText strChefCode, adWriteLine
                End If
                ActiveRow = ActiveRow + 1
            Loop
        Case Else
    End Select
End Function


'------------------------------'
'フッター作成
'------------------------------'
Private Function CreateFooter(ByVal strFileType As String)
    'ファイル種類判定
    Select Case strFileType
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            WriteStream.WriteText indent2
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        Case Else
    End Select
End Function


'------------------------------'
'パターンAコード生成
'------------------------------'
Private Function CreatePatternA(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute名の列を取得
    AttributeColumn = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                   
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
                            
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
        Case Else
    End Select
End Function

'------------------------------'
'パターンBコード生成
'------------------------------'
Private Function CreatePatternB(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute名取得
    AttributeRow = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            'Attribute名書き出し
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            'Attribute名書き出し
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
         Case Else
    End Select
End Function

'------------------------------'
'パターンCコード生成
'------------------------------'
Private Function CreatePatternC(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("以上（追加項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute名の列を取得
    AttributeColumn = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "以上（変更項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing And Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '------------------------------'
            '追加項目処理
            '------------------------------'
            '追加項目処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="○")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute名書き出し
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
                            
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "以上（変更項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing And Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '------------------------------'
            '追加項目処理
            '------------------------------'
            '追加項目処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="○")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute名書き出し
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        Case Else
    End Select
End Function

'------------------------------'
'パターンDコード生成（grub.conf）
'------------------------------'
Private Function CreatePatternD(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上（変更項目）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("以上（追加項目(kernel行)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow3 = WS.Cells.Find("以上（追加項目(ファイル末尾)）", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute名の列を取得
    AttributeColumn = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "以上（変更項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing And Rng3 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '------------------------------'
            '追加項目(kernel行)処理
            '------------------------------'
            '追加項目(kernel行)処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
            If Not Rng2 Is Nothing Then
                'Attribute名書き出し
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '追加項目(kernel行)列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目(kernel行)）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '追加項目(ファイル末尾)処理
            '------------------------------'
            '追加項目(ファイル末尾)処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
            If Not Rng3 Is Nothing Then
                'Attribute名書き出し
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '追加項目(ファイル末尾)列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目(ファイル末尾)）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If

        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            '------------------------------'
            '変更項目処理
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "以上（変更項目）"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute名・設定値取得、書き出し&エスケープ文字付加
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    'カンマ判定処理（最終処理）
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
                    'アクティブシートが最終処理シートかつアクティブセル以下に値がない場合
                    If strTargetSheet = strLastSheet And Rng Is Nothing And Rng3 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                        Exit Function
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '------------------------------'
            '追加項目(kernel行)処理
            '------------------------------'
            '追加項目(kernel行)処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
            If Not Rng2 Is Nothing Then
                'Attribute名書き出し
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '追加項目(kernel行)列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目(kernel行)）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    'コード作成＆エスケープ文字付加
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '追加項目(ファイル末尾)処理
            '------------------------------'
            '追加項目(ファイル末尾)処理開始行設定
            ActiveRow = ActiveRow + 2
            
            'Attribute書き出し判定
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="○")
            If Not Rng3 Is Nothing Then
                'Attribute名書き出し
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '追加項目(ファイル末尾)列ループ処理
            Do While WS.Cells(ActiveRow, 12) <> "以上（追加項目(ファイル末尾)）"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        Case Else
    End Select
End Function


'------------------------------'
'パターンEコード生成
'------------------------------'
Private Function CreatePatternE(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '変数宣言
    Dim HashRow As Long
    Dim HashColumn As Long
    Dim HashCount As Long
    Dim strHashValue As String
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Attribute名取得
    AttributeRow = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute名", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    'Hash個数取得
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash名位置取得
    HashRow = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            'Attribute名書き出し
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    strChefCode = indent4 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '初期化
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash数分行ループ処理
                    Do While Count < HashCount
                        'ハッシュキー、ハッシュ値取得　書き出し&エスケープ文字付加
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn).Value
                        strHashValue = AddEscape(strHashValue)
    '★
                        strChefCode = indent5 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        'カンマ判定処理（Hash内）
                        If Count = HashCount - 1 Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                        
                        Count = Count + 1
                        ActiveColumn = ActiveColumn + 1
                        HashColumn = HashColumn + 1
                    Loop
                    
                    WriteStream.WriteText indent4
                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            'Attribute名書き出し
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '追加項目列ループ処理
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "○" Then
                    strChefCode = indent3 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '初期化
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash名", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash数分行ループ処理
                    Do While Count < HashCount
                        'ハッシュキー、ハッシュ値取得　書き出し&エスケープ文字付加
                        strHashValue = AddEscape(WS.Cells(ActiveRow, ActiveColumn).Value)
                        strChefCode = indent4 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        'カンマ判定処理（Hash内）
                        If Count = HashCount - 1 Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                        
                        Count = Count + 1
                        ActiveColumn = ActiveColumn + 1
                        HashColumn = HashColumn + 1
                    Loop
                    
                    WriteStream.WriteText indent3
                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="○")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            'カンマ判定処理（最終処理）
            If strTargetSheet = strLastSheet Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        Case Else
    End Select
End Function

'------------------------------'
'パターンFコード生成（自動起動サービス）
'------------------------------'
Private Function CreatePatternF(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '変数宣言
    Dim strRunLevel As String
    Dim TemplateColumn As Long
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Templateの列を設定
    TemplateColumn = 7
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
        
            '------------------------------'
            'Service Addコード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent3 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'カンマ判定処理（Hash内）
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "以上"
                        If WS.Cells(Line, TemplateColumn).Value = "-" And ((Mid(WS.Cells(Line, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(Line, ActiveFileColumn).Value, 3) = "on")) Then
                            HashWrite = 1
                        End If
                        Line = Line + 1
                    Loop
                    If HashWrite = 1 Then
                        strChefCode = strChefCode + ","
                    End If
                
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strChefCode, adWriteLine
                
                End If
                ActiveRow = ActiveRow + 1
            Loop
            If AttWrite = 1 Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Service Deleteコード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent3 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'カンマ判定処理（Hash内）
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "以上"
                        If WS.Cells(Line, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(Line, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(Line, TemplateColumn).Value, 3) = "on")) Then
                            HashWrite = 1
                        End If
                        Line = Line + 1
                    Loop
                    If HashWrite = 1 Then
                        strChefCode = strChefCode + ","
                    End If
                
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strChefCode, adWriteLine
                
                End If
                ActiveRow = ActiveRow + 1
            Loop
            If AttWrite = 1 Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Service Change(on)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                'Attribute書き出し判定
                If AttWrite = 0 Then
                    'Attribute名書き出し
                    strChefCode = indent3 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ランレベルの個数分ループ処理
                Count = 0
                strRunLevel = ""
                Do While Count < 7
                    If Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 3) = "on" Then
                        strRunLevel = strRunLevel + Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 1, 1)
                    End If
            
                    Count = Count + 1
                Loop

                If strRunLevel <> "" Then
                    strRunLevel = Dquote + "level" + Dquote + ": " + Dquote + strRunLevel + Dquote
                    strChefCode = Dquote + "name" + Dquote + ": " + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote + ","
                
                    WriteStream.WriteText indent4
                    WriteStream.WriteText "{", adWriteLine
                    WriteStream.WriteText indent5
                    WriteStream.WriteText strChefCode, adWriteLine
                    WriteStream.WriteText indent5
                    WriteStream.WriteText strRunLevel, adWriteLine
                    WriteStream.WriteText indent4

                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn + 6).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="?:on")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
             
                End If
                ActiveRow = ActiveRow + 1
            Loop
            If AttWrite = 1 Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "],", adWriteLine
            End If

            '------------------------------'
            'Service Change(off)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                'Attribute書き出し判定
                If AttWrite = 0 Then
                    'Attribute名書き出し
                    strChefCode = indent3 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ランレベルの個数分ループ処理
                Count = 0
                strRunLevel = ""
                Do While Count < 7
                    If Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 3) = "off" Then
                        strRunLevel = strRunLevel + Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 1, 1)
                    End If
            
                    Count = Count + 1
                Loop

                If strRunLevel <> "" Then
                    strRunLevel = Dquote + "level" + Dquote + ": " + Dquote + strRunLevel + Dquote
                    strChefCode = Dquote + "name" + Dquote + ": " + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote + ","
                
                    WriteStream.WriteText indent4
                    WriteStream.WriteText "{", adWriteLine
                    WriteStream.WriteText indent5
                    WriteStream.WriteText strChefCode, adWriteLine
                    WriteStream.WriteText indent5
                    WriteStream.WriteText strRunLevel, adWriteLine
                    WriteStream.WriteText indent4

                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn + 6).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="?:off")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
             
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            If AttWrite = 1 Then
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    strChefCode = indent3 + "]"
                    WriteStream.WriteText strChefCode, adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "],", adWriteLine
                End If
            End If
            
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
        
            '------------------------------'
            'Service Addコード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent2 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'カンマ判定処理（Hash内）
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "以上"
                        If WS.Cells(Line, TemplateColumn).Value = "-" And ((Mid(WS.Cells(Line, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(Line, ActiveFileColumn).Value, 3) = "on")) Then
                            HashWrite = 1
                        End If
                        Line = Line + 1
                    Loop
                    If HashWrite = 1 Then
                        strChefCode = strChefCode + ","
                    End If
                
                    WriteStream.WriteText indent3
                    WriteStream.WriteText strChefCode, adWriteLine
                
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            If AttWrite = 1 Then
                WriteStream.WriteText indent2
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Service Deleteコード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent2 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'カンマ判定処理（Hash内）
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "以上"
                        If WS.Cells(Line, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(Line, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(Line, TemplateColumn).Value, 3) = "on")) Then
                            HashWrite = 1
                        End If
                        Line = Line + 1
                    Loop
                    If HashWrite = 1 Then
                        strChefCode = strChefCode + ","
                    End If
                
                    WriteStream.WriteText indent3
                    WriteStream.WriteText strChefCode, adWriteLine
                
                End If
                ActiveRow = ActiveRow + 1
            Loop

            If AttWrite = 1 Then
                WriteStream.WriteText indent2
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Service Change(on)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                'Attribute書き出し判定
                If AttWrite = 0 Then
                    'Attribute名書き出し
                    strChefCode = indent2 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ランレベルの個数分ループ処理
                Count = 0
                strRunLevel = ""
                Do While Count < 7
                    If Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 3) = "on" Then
                        strRunLevel = strRunLevel + Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 1, 1)
                    End If
            
                    Count = Count + 1
                Loop

                If strRunLevel <> "" Then
                    strRunLevel = Dquote + "level" + Dquote + ": " + Dquote + strRunLevel + Dquote
                    strChefCode = Dquote + "name" + Dquote + ": " + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote + ","
                
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "{", adWriteLine
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strChefCode, adWriteLine
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strRunLevel, adWriteLine
                    WriteStream.WriteText indent3

                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn + 6).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="?:on")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
             
                End If
                ActiveRow = ActiveRow + 1
            Loop

            If AttWrite = 1 Then
                WriteStream.WriteText indent2
                WriteStream.WriteText "],", adWriteLine
            End If

            '------------------------------'
            'Service Change(off)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                'Attribute書き出し判定
                If AttWrite = 0 Then
                    'Attribute名書き出し
                    strChefCode = indent2 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ランレベルの個数分ループ処理
                Count = 0
                strRunLevel = ""
                Do While Count < 7
                    If Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 3) = "off" Then
                        strRunLevel = strRunLevel + Mid(WS.Cells(ActiveRow, ActiveFileColumn + Count).Value, 1, 1)
                    End If
            
                    Count = Count + 1
                Loop

                If strRunLevel <> "" Then
                    strRunLevel = Dquote + "level" + Dquote + ": " + Dquote + strRunLevel + Dquote
                    strChefCode = Dquote + "name" + Dquote + ": " + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote + ","
                
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "{", adWriteLine
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strChefCode, adWriteLine
                    WriteStream.WriteText indent4
                    WriteStream.WriteText strRunLevel, adWriteLine
                    WriteStream.WriteText indent3

                    'カンマ判定処理（Hash外）
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn + 6).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="?:off")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
             
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            If AttWrite = 1 Then
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    strChefCode = indent2 + "]"
                    WriteStream.WriteText strChefCode, adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent2
                    WriteStream.WriteText "],", adWriteLine
                End If
            End If
        Case Else
    End Select
End Function

'------------------------------'
'パターンGコード生成（xinetdサービス）
'------------------------------'
Private Function CreatePatternG(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '開始行設定
    ActiveRow = 6
    
    '終了行設定
    EndRow = WS.Cells.Find("以上", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            '------------------------------'
            'Xinetd Service Change(on)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent3 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                        strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                        Set Rng = WS.Range(strSerchTarget).Find(what:="on")
                        If Rng Is Nothing Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                End If
                ActiveRow = ActiveRow + 1
            Loop

            If AttWrite = 1 Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Xinetd Service Change(off)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent3 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                        strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                        Set Rng = WS.Range(strSerchTarget).Find(what:="off")
                        If Rng Is Nothing Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            If AttWrite = 1 Then
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    strChefCode = indent3 + "]"
                    WriteStream.WriteText strChefCode, adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "],", adWriteLine
                End If
            End If

        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            '------------------------------'
            'Xinetd Service Change(on)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent2 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                        strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                        Set Rng = WS.Range(strSerchTarget).Find(what:="on")
                        If Rng Is Nothing Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                End If
                ActiveRow = ActiveRow + 1
            Loop

            If AttWrite = 1 Then
                WriteStream.WriteText indent2
                WriteStream.WriteText "],", adWriteLine
            End If
            
            '------------------------------'
            'Xinetd Service Change(off)コード作成
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "以上"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute書き出し判定
                    If AttWrite = 0 Then
                        'Attribute名書き出し
                        strChefCode = indent2 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'カンマ判定処理（配列内）
                        strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                        Set Rng = WS.Range(strSerchTarget).Find(what:="off")
                        If Rng Is Nothing Then
                            WriteStream.WriteText "", adWriteLine
                        Else
                            WriteStream.WriteText ",", adWriteLine
                        End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            If AttWrite = 1 Then
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    strChefCode = indent2 + "]"
                    WriteStream.WriteText strChefCode, adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent2
                    WriteStream.WriteText "],", adWriteLine
                End If
            End If

        Case Else
    End Select
End Function

'------------------------------'
'パターンHコード生成（hosts）
'------------------------------'
Private Function CreatePatternH(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'Attribute名設定
    strAttributeName = "att_ho_hosts_chg_server"
    
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        Case Else
    End Select
        
    'カンマ判定処理（最終処理）
    If strTargetSheet = strLastSheet Then
        WriteStream.WriteText "", adWriteLine
        Exit Function
    Else
        WriteStream.WriteText ",", adWriteLine
    End If

End Function





'------------------------------'
'パターンIコード生成（ユーザ環境変数）
'------------------------------'
Private Function CreatePatternI(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '変数宣言
    Dim strLastEvSheet As String
    Dim strFirstEvSheet As String
    
    'ワークシートオブジェクト設定
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
    Set FlagWS = ThisWorkbook.Worksheets("処理シート")
     
    '対象ファイルの列を設定
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'Attribute名設定
    strAttributeName = "att_ev_chg_environment_variable"
    
    '開始・終了のユーザ環境変数シート名を取得
    AttWriteI = 0
    ActiveRow = 4
    Do While FlagWS.Cells(ActiveRow, 3) <> ""
        If FlagWS.Cells(ActiveRow, 3).Value = "I" Then
            If AttWriteI = 0 Then
                strFirstEvSheet = FlagWS.Cells(ActiveRow, 2).Value
                AttWriteI = 1
            ElseIf AttWriteI = 1 Then
                strLastEvSheet = FlagWS.Cells(ActiveRow, 2).Value
            End If
        End If
        ActiveRow = ActiveRow + 1
    Loop
  
    'ファイル種類判定
    Select Case strFileType
    
        '------------------------------'
        'Roleファイル処理
        '------------------------------'
        Case "Role"
            'Attribute書き出し判定
            If strTargetSheet = strFirstEvSheet Then
                'Attribute名書き出し
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            strChefCode = indent4 + "{"
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent5 + Dquote + "server" + Dquote + ": " + Dquote + strFileName + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent5 + Dquote + "user_name" + Dquote + ": " + Dquote + WS.Cells(6, 3).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent5 + Dquote + "group_name" + Dquote + ": " + Dquote + WS.Cells(6, 4).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent5 + Dquote + "path" + Dquote + ": " + Dquote + WS.Cells(6, 5).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent5 + Dquote + "file_name" + Dquote + ": " + Dquote + WS.Cells(6, 6).Value + Dquote
            WriteStream.WriteText strChefCode, adWriteLine
            
            'カンマ判定処理（Hash外）
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent4
                WriteStream.WriteText "}", adWriteLine
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "]", adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent3
                    WriteStream.WriteText "],", adWriteLine
                End If
            Else
                WriteStream.WriteText indent4
                WriteStream.WriteText "},", adWriteLine
            End If
                  
        '------------------------------'
        'Nodeファイル処理
        '------------------------------'
        Case "Node"
            'Attribute書き出し判定
            If strTargetSheet = strFirstEvSheet Then
                'Attribute名書き出し
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            strChefCode = indent3 + "{"
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent4 + Dquote + "server" + Dquote + ": " + Dquote + strFileName + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent4 + Dquote + "user_name" + Dquote + ": " + Dquote + WS.Cells(6, 3).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent4 + Dquote + "group_name" + Dquote + ": " + Dquote + WS.Cells(6, 4).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent4 + Dquote + "path" + Dquote + ": " + Dquote + WS.Cells(6, 5).Value + Dquote + ","
            WriteStream.WriteText strChefCode, adWriteLine
            strChefCode = indent4 + Dquote + "file_name" + Dquote + ": " + Dquote + WS.Cells(6, 6).Value + Dquote
            WriteStream.WriteText strChefCode, adWriteLine
            
            'カンマ判定処理（Hash外）
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "}", adWriteLine
                'カンマ判定処理（最終処理）
                If strTargetSheet = strLastSheet Then
                    WriteStream.WriteText indent4
                    WriteStream.WriteText "]", adWriteLine
                    Exit Function
                Else
                    WriteStream.WriteText indent4
                    WriteStream.WriteText "],", adWriteLine
                End If
            Else
                WriteStream.WriteText indent3
                WriteStream.WriteText "},", adWriteLine
            End If
            
        Case Else
    End Select


End Function



