Sub Sample1()
    Dim i As Long, j As Long, cnt As Long
    Dim buf() As String, swap As String
    
    cnt = Worksheets.Count
    ReDim buf(cnt)
    
    'ワークシート名を配列に入れる
    For i = 1 To cnt
        buf(i) = Worksheets(i).Name
    Next i
    
    '配列の要素をソートする
    For i = 1 To cnt
        For j = cnt To i Step -1
            If buf(i) > buf(j) Then
                swap = buf(i)
                buf(i) = buf(j)
                buf(j) = swap
            End If
        Next j
    Next i
    
    'ワークシートの位置を並べ替える
    Worksheets(buf(1)).Move Before:=Worksheets(1)
    For i = 2 To cnt
        Worksheets(buf(i)).Move After:=Worksheets(i - 1)
    Next i
End Sub
