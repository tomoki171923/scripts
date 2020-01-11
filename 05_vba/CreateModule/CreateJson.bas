Attribute VB_Name = "CreateJson"
'Option Explicit
    '�ϐ��錾�i�O���[�o���j
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
    
    '�萔�錾�i�O���[�o���j
    Public Const Dquote As String = """"
    Public Const indent1 As String = "  "
    Public Const indent2 As String = "    "
    Public Const indent3 As String = "      "
    Public Const indent4 As String = "        "
    Public Const indent5 As String = "          "
    
'------------------------------'
'�R���g���[���[
'------------------------------'
Sub ControllerJson()
    '�ϐ��錾
    Dim StartColumn As Long
    Dim StartRow As Long
    
    '�t�H���_����
    OutputPath = ThisWorkbook.Worksheets("���C��").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "�o�͐�t�H���_�Ɍ�肪����܂��B", vbCritical
        Exit Sub
    End If

    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets("�R�[�h�ꗗ")

    '�J�n�ʒu�ݒ�
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
        '������
        StartColumn = 3
        strNodeName = ""
        
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'�p�^�[������
'------------------------------'
Private Function JudgePattern(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iOpen�j
    Call OpenTextStream
    
    '�w�b�_�[���쐬����
    Call CreateHeader(strFileType, strRoleName, strNodeName)
    
    '�t�@�C�����ݒ�
    Select Case strFileType
        Case "Role"
            strFileName = strRoleName
        Case "Node"
            strFileName = strNodeName
        Case Else
    End Select
       
    '���[�N�V �[�g�I�u�W�F�N�g�ݒ聕�N���A
    Set FlagWS = ThisWorkbook.Worksheets("�����V�[�g")
    FlagWS.Range("B4:D999999").ClearContents
       
    '�����V�[�g�ݒ�
    Call JudgeSheets(strFileName)

    '�ŏI�����V�[�g���擾
    strLastSheet = FlagWS.Range("B3").End(xlDown).Value
        
    a = 4
    Do While FlagWS.Cells(a, 2) <> ""
        strTargetSheet = FlagWS.Cells(a, 2)
        strPatternType = FlagWS.Cells(a, 3)
    
        '�p�^�[������
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
    
    '�t�b�^�[���쐬����
    Call CreateFooter(strFileType)

    ' �����o���֐��Ăяo��
    strFileName = strFileName + ".json"
    Call FileOutput(strFileName)

    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iClose�j
    Call CloseTextStream

End Function

'------------------------------'
'�����V�[�g����
'------------------------------'
Private Function JudgeSheets(ByVal strFileName As String) As String
    '�ϐ��錾
    Dim PatternCell As Range
    i = 4
    '�V�[�g���[�v
    For Each TargetSheet In ThisWorkbook.Worksheets
        '�p�^�[��������
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        If PatternCell Is Nothing Then
            strPatternType = ""
        Else
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
        End If
    
        '�J�n�s�ݒ�
        ActiveRow = 6
 
        '�Ώۃt�@�C���̗�ɐݒ�l���܂߂�Ă��邩����
        Select Case strPatternType
            Case "A"
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
                
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
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
        
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="��")
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
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�i�ǉ����ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '�Ώۃt�@�C���̗��"��"���܂܂�Ă��邩����
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="��")

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
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�i�ǉ�����(kernel�s)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '�Ώۃt�@�C���̗��"��"���܂܂�Ă��邩����
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="��")

                If Rng Is Nothing And strFlag = "off" Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�i�ǉ�����(�t�@�C������)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '�Ώۃt�@�C���̗��"��"���܂܂�Ă��邩����
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="��")

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
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row

                'Template�̗��ݒ�
                TemplateColumn = 7
    
                '�Ώۃt�@�C���̗�ɐݒ�l���܂܂�Ă��邩����
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
                '�Ώۃt�@�C���̗��ݒ�
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                '�I���s�ݒ�
                EndRow = TargetSheet.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
    
                '�Ώۃt�@�C���̗�ɐݒ�l���܂܂�Ă��邩����
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
'�G�X�P�[�v�����t��
'------------------------------'
Private Function AddEscape(ByVal strTarget As String) As String
    '�_�u���N�H�[�e�[�V����������ꍇ�G�X�P�[�v�����ǉ�
    AddEscape = Replace(strTarget, """", "�""")
End Function


'------------------------------'
'�w�b�_�[�쐬
'------------------------------'
Private Function CreateHeader(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
        '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets("���V�s�ꗗ")
    
    '�J�n�s�ݒ�
    ActiveRow = 4
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
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
        
        
            'CookBook���擾
            Do While ActiveRow < EndRow
                If WS.Cells(ActiveRow, 3).Value <> "" Then
                    If WS.Cells(ActiveRow + 1, 3).Value <> "" Then
                        '�����s���ݒ�
                        Line = 1
                    Else
                        '�����s���ݒ�
                        LineCount = WS.Cells(ActiveRow, 3).End(xlDown).Row
                        Line = LineCount - ActiveRow
                    End If
                    strCookbookName = WS.Cells(ActiveRow, 3).Value
                
                    j = 1
                    'Recipe���擾
                    Do While j <= Line
                        strRecipeName = WS.Cells(ActiveRow, 4).Value
                        
                        If strRecipeName <> "" Then
                            '�J���}���菈���irun_list���j
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
        'Node�t�@�C������
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
            
            'CookBook���擾
            Do While WS.Cells(ActiveRow, 3).Value <> "�ȏ�"
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
'�t�b�^�[�쐬
'------------------------------'
Private Function CreateFooter(ByVal strFileType As String)
    '�t�@�C����ޔ���
    Select Case strFileType
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            WriteStream.WriteText indent2
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        Case Else
    End Select
End Function


'------------------------------'
'�p�^�[��A�R�[�h����
'------------------------------'
Private Function CreatePatternA(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute���̗���擾
    AttributeColumn = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                   
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
'�p�^�[��B�R�[�h����
'------------------------------'
Private Function CreatePatternB(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute���擾
    AttributeRow = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            'Attribute�������o��
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            'Attribute�������o��
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
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
'�p�^�[��C�R�[�h����
'------------------------------'
Private Function CreatePatternC(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("�ȏ�i�ǉ����ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute���̗���擾
    AttributeColumn = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ύX���ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
            '�ǉ����ڏ���
            '------------------------------'
            '�ǉ����ڏ����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="��")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute�������o��
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ����ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
                            
        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ύX���ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
            '�ǉ����ڏ���
            '------------------------------'
            '�ǉ����ڏ����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="��")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute�������o��
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ����ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
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
'�p�^�[��D�R�[�h�����igrub.conf�j
'------------------------------'
Private Function CreatePatternD(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("�ȏ�i�ǉ�����(kernel�s)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow3 = WS.Cells.Find("�ȏ�i�ǉ�����(�t�@�C������)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute���̗���擾
    AttributeColumn = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "�ȏ�i�ύX���ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
            '�ǉ�����(kernel�s)����
            '------------------------------'
            '�ǉ�����(kernel�s)�����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
            If Not Rng2 Is Nothing Then
                'Attribute�������o��
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '�ǉ�����(kernel�s)�񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ�����(kernel�s)�j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '�ǉ�����(�t�@�C������)����
            '------------------------------'
            '�ǉ�����(�t�@�C������)�����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
            If Not Rng3 Is Nothing Then
                'Attribute�������o��
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '�ǉ�����(�t�@�C������)�񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ�����(�t�@�C������)�j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If

        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            '------------------------------'
            '�ύX���ڏ���
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "�ȏ�i�ύX���ځj"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute���E�ݒ�l�擾�A�����o��&�G�X�P�[�v�����t��
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    '�J���}���菈���i�ŏI�����j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
                    '�A�N�e�B�u�V�[�g���ŏI�����V�[�g���A�N�e�B�u�Z���ȉ��ɒl���Ȃ��ꍇ
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
            '�ǉ�����(kernel�s)����
            '------------------------------'
            '�ǉ�����(kernel�s)�����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
            If Not Rng2 Is Nothing Then
                'Attribute�������o��
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '�ǉ�����(kernel�s)�񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ�����(kernel�s)�j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    '�R�[�h�쐬���G�X�P�[�v�����t��
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '�ǉ�����(�t�@�C������)����
            '------------------------------'
            '�ǉ�����(�t�@�C������)�����J�n�s�ݒ�
            ActiveRow = ActiveRow + 2
            
            'Attribute�����o������
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="��")
            If Not Rng3 Is Nothing Then
                'Attribute�������o��
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '�ǉ�����(�t�@�C������)�񃋁[�v����
            Do While WS.Cells(ActiveRow, 12) <> "�ȏ�i�ǉ�����(�t�@�C������)�j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            '�J���}���菈���i�ŏI�����j
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
'�p�^�[��E�R�[�h����
'------------------------------'
Private Function CreatePatternE(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '�ϐ��錾
    Dim HashRow As Long
    Dim HashColumn As Long
    Dim HashCount As Long
    Dim strHashValue As String
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Attribute���擾
    AttributeRow = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute��", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    'Hash���擾
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash���ʒu�擾
    HashRow = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            'Attribute�������o��
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    strChefCode = indent4 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '������
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash�����s���[�v����
                    Do While Count < HashCount
                        '�n�b�V���L�[�A�n�b�V���l�擾�@�����o��&�G�X�P�[�v�����t��
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn).Value
                        strHashValue = AddEscape(strHashValue)
    '��
                        strChefCode = indent5 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        '�J���}���菈���iHash���j
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
                    '�J���}���菈���iHash�O�j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            '�J���}���菈���i�ŏI�����j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        
        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            'Attribute�������o��
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '�ǉ����ڗ񃋁[�v����
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
                    strChefCode = indent3 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '������
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash�����s���[�v����
                    Do While Count < HashCount
                        '�n�b�V���L�[�A�n�b�V���l�擾�@�����o��&�G�X�P�[�v�����t��
                        strHashValue = AddEscape(WS.Cells(ActiveRow, ActiveColumn).Value)
                        strChefCode = indent4 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        '�J���}���菈���iHash���j
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
                    '�J���}���菈���iHash�O�j
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="��")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            '�J���}���菈���i�ŏI�����j
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
'�p�^�[��F�R�[�h�����i�����N���T�[�r�X�j
'------------------------------'
Private Function CreatePatternF(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '�ϐ��錾
    Dim strRunLevel As String
    Dim TemplateColumn As Long
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Template�̗��ݒ�
    TemplateColumn = 7
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
        
            '------------------------------'
            'Service Add�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent3 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    '�J���}���菈���iHash���j
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "�ȏ�"
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
            'Service Delete�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent3 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    '�J���}���菈���iHash���j
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "�ȏ�"
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
            'Service Change(on)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                'Attribute�����o������
                If AttWrite = 0 Then
                    'Attribute�������o��
                    strChefCode = indent3 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                '�������x���̌������[�v����
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

                    '�J���}���菈���iHash�O�j
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
            'Service Change(off)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                'Attribute�����o������
                If AttWrite = 0 Then
                    'Attribute�������o��
                    strChefCode = indent3 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                '�������x���̌������[�v����
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

                    '�J���}���菈���iHash�O�j
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
                '�J���}���菈���i�ŏI�����j
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
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
        
            '------------------------------'
            'Service Add�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent2 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    '�J���}���菈���iHash���j
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "�ȏ�"
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
            'Service Delete�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent2 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    '�J���}���菈���iHash���j
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "�ȏ�"
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
            'Service Change(on)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                'Attribute�����o������
                If AttWrite = 0 Then
                    'Attribute�������o��
                    strChefCode = indent2 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                '�������x���̌������[�v����
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

                    '�J���}���菈���iHash�O�j
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
            'Service Change(off)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                'Attribute�����o������
                If AttWrite = 0 Then
                    'Attribute�������o��
                    strChefCode = indent2 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                '�������x���̌������[�v����
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

                    '�J���}���菈���iHash�O�j
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
                '�J���}���菈���i�ŏI�����j
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
'�p�^�[��G�R�[�h�����ixinetd�T�[�r�X�j
'------------------------------'
Private Function CreatePatternG(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            '------------------------------'
            'Xinetd Service Change(on)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent3 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
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
            'Xinetd Service Change(off)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent3 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
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
                '�J���}���菈���i�ŏI�����j
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
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            '------------------------------'
            'Xinetd Service Change(on)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent2 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
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
            'Xinetd Service Change(off)�R�[�h�쐬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute�����o������
                    If AttWrite = 0 Then
                        'Attribute�������o��
                        strChefCode = indent2 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    '�J���}���菈���i�z����j
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
                '�J���}���菈���i�ŏI�����j
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
'�p�^�[��H�R�[�h�����ihosts�j
'------------------------------'
Private Function CreatePatternH(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'Attribute���ݒ�
    strAttributeName = "att_ho_hosts_chg_server"
    
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        '------------------------------'
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        Case Else
    End Select
        
    '�J���}���菈���i�ŏI�����j
    If strTargetSheet = strLastSheet Then
        WriteStream.WriteText "", adWriteLine
        Exit Function
    Else
        WriteStream.WriteText ",", adWriteLine
    End If

End Function





'------------------------------'
'�p�^�[��I�R�[�h�����i���[�U���ϐ��j
'------------------------------'
Private Function CreatePatternI(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '�ϐ��錾
    Dim strLastEvSheet As String
    Dim strFirstEvSheet As String
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
    Set FlagWS = ThisWorkbook.Worksheets("�����V�[�g")
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'Attribute���ݒ�
    strAttributeName = "att_ev_chg_environment_variable"
    
    '�J�n�E�I���̃��[�U���ϐ��V�[�g�����擾
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
  
    '�t�@�C����ޔ���
    Select Case strFileType
    
        '------------------------------'
        'Role�t�@�C������
        '------------------------------'
        Case "Role"
            'Attribute�����o������
            If strTargetSheet = strFirstEvSheet Then
                'Attribute�������o��
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
            
            '�J���}���菈���iHash�O�j
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent4
                WriteStream.WriteText "}", adWriteLine
                '�J���}���菈���i�ŏI�����j
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
        'Node�t�@�C������
        '------------------------------'
        Case "Node"
            'Attribute�����o������
            If strTargetSheet = strFirstEvSheet Then
                'Attribute�������o��
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
            
            '�J���}���菈���iHash�O�j
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "}", adWriteLine
                '�J���}���菈���i�ŏI�����j
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



