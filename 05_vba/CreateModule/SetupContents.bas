Attribute VB_Name = "SetupContents"
'------------------------------'
'�ڎ��X�V
'------------------------------'
Sub ContentsUpdate()
    Dim i As Long
    Dim strLinkPath As String
    
    Set WS = ThisWorkbook.Worksheets("�ڎ�")
    WS.Cells.Borders.LineStyle = xlLineStyleNone
    WS.Range("B3:D999999").ClearContents
    
    For i = 3 To Sheets.Count
    
        '���ԂƃV�[�g�������o��
        WS.Cells(i, 2).Value = i - 2
        WS.Cells(i, 3).Value = Worksheets(i).Name
        
        '�����N�쐬
        strLinkPath = "#'" & Worksheets(i).Name & "'!A1"
        WS.Hyperlinks.Add Anchor:=WS.Cells(i, 3), Address:=strLinkPath
        
        couban = couban + 1
    Next i

    WS.Range("B2", Cells(i - 1, 3)).Borders.LineStyle = xlContinuous
    
End Sub

