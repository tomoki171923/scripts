Attribute VB_Name = "���̑�"
Sub �s��폜�}�N��()
    
    '�ΏۃV�[�g�Ɋ܂܂�Ă��镶������w��
    strSerchSheet = "�y"
    '�Ώۗ�Ɋ܂܂�Ă��镶������w��
    strSerchColumn = ""
    '�Ώۍs�Ɋ܂܂�Ă��镶������w��
    strSerchRow = ""

    '�S�V�[�g���[�v
    For Each TargetSheet In ThisWorkbook.Worksheets
       ' '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
        'Set WS = ThisWorkbook.Worksheets(TargetSheet.Name)
        
        '�Ώۂ̗������
        'TargetColumn = WS.Cells.Find(strSerchColumn, LookIn:=xlValues, LookAt:=xlWhole).Column
        
        '�Ώۂ̍s������
        'TargetRow = WS.Cells.Find(strSerchRow, LookIn:=xlValues, LookAt:=xlWhole).Row
        
        '�w�肵�������񂪃V�[�g�Ɋ܂܂�Ă��邩����
        If InStr(TargetSheet.Name, SerchSheetName) > 0 Then
            '�w�肵�������񂪊܂܂�Ă������폜
                '�s���폜���܂�
            Range("T:T").Delete
            '�w�肵�������񂪊܂܂�Ă���s���폜
            'Rows(TargetRow).Delete
        End If
    Next

End Sub

