Sub Sample1()
    Dim i As Long, j As Long, cnt As Long
    Dim buf() As String, swap As String
    
    cnt = Worksheets.Count
    ReDim buf(cnt)
    
    '���[�N�V�[�g����z��ɓ����
    For i = 1 To cnt
        buf(i) = Worksheets(i).Name
    Next i
    
    '�z��̗v�f���\�[�g����
    For i = 1 To cnt
        For j = cnt To i Step -1
            If buf(i) > buf(j) Then
                swap = buf(i)
                buf(i) = buf(j)
                buf(j) = swap
            End If
        Next j
    Next i
    
    '���[�N�V�[�g�̈ʒu����בւ���
    Worksheets(buf(1)).Move Before:=Worksheets(1)
    For i = 2 To cnt
        Worksheets(buf(i)).Move After:=Worksheets(i - 1)
    Next i
End Sub
