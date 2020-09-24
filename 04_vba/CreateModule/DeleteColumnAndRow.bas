Attribute VB_Name = "その他"
Sub 行列削除マクロ()
    
    '対象シートに含まれている文字列を指定
    strSerchSheet = "【"
    '対象列に含まれている文字列を指定
    strSerchColumn = ""
    '対象行に含まれている文字列を指定
    strSerchRow = ""

    '全シートループ
    For Each TargetSheet In ThisWorkbook.Worksheets
       ' 'ワークシートオブジェクト設定
        'Set WS = ThisWorkbook.Worksheets(TargetSheet.Name)
        
        '対象の列を検索
        'TargetColumn = WS.Cells.Find(strSerchColumn, LookIn:=xlValues, LookAt:=xlWhole).Column
        
        '対象の行を検索
        'TargetRow = WS.Cells.Find(strSerchRow, LookIn:=xlValues, LookAt:=xlWhole).Row
        
        '指定した文字列がシートに含まれているか判定
        If InStr(TargetSheet.Name, SerchSheetName) > 0 Then
            '指定した文字列が含まれている列を削除
                '行を削除します
            Range("T:T").Delete
            '指定した文字列が含まれている行を削除
            'Rows(TargetRow).Delete
        End If
    Next

End Sub

