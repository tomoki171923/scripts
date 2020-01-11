Attribute VB_Name = "SetupContents"
'------------------------------'
'目次更新
'------------------------------'
Sub ContentsUpdate()
    Dim i As Long
    Dim strLinkPath As String
    
    Set WS = ThisWorkbook.Worksheets("目次")
    WS.Cells.Borders.LineStyle = xlLineStyleNone
    WS.Range("B3:D999999").ClearContents
    
    For i = 3 To Sheets.Count
    
        '項番とシート名書き出し
        WS.Cells(i, 2).Value = i - 2
        WS.Cells(i, 3).Value = Worksheets(i).Name
        
        'リンク作成
        strLinkPath = "#'" & Worksheets(i).Name & "'!A1"
        WS.Hyperlinks.Add Anchor:=WS.Cells(i, 3), Address:=strLinkPath
        
        couban = couban + 1
    Next i

    WS.Range("B2", Cells(i - 1, 3)).Borders.LineStyle = xlContinuous
    
End Sub

