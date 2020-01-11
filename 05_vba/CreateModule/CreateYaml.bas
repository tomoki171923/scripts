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
'�R���g���[��
'------------------------------'
Sub ControllerYaml()
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
    Set ListWS = ThisWorkbook.Worksheets("�R�[�h�ꗗ")

    'TargetList�t�@�C���쐬
    Call CreateTargetList

    '�J�n�ʒu�ݒ�
    StartColumn = 3
    StartRow = 5
    
    '�eyaml�t�@�C���쐬
    Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
        strFileName = ListWS.Cells(StartRow, StartColumn).Value
        Call JudgePattern(strFileName)
            
        StartColumn = StartColumn + 1
        Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
            strFileName = ListWS.Cells(StartRow, StartColumn).Value
            Call JudgePattern(strFileName)
            StartColumn = StartColumn + 1
        Loop
        '������
        StartColumn = 3
        strFileName = ""
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'�p�^�[������
'------------------------------'
Private Function JudgePattern(ByVal strFileName As String)
    Dim i As Long
    Dim j As Long
    Dim PatternCell As Object
    
    '���[�N�V �[�g�I�u�W�F�N�g�ݒ聕�N���A
    Set FlagWS = ThisWorkbook.Worksheets("�����V�[�g")
    FlagWS.Range("B4:D999999").ClearContents

    i = 4
    '�����V�[�g�ɏ����o��
    For Each TargetSheet In ThisWorkbook.Worksheets
        '�p�^�[��������
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not PatternCell Is Nothing Then
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
            
            '�Ώۃt�@�C����ɕ����񂪊܂�ł��邩����
            If JudgeSheet(strFileName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = strPatternType
                i = i + 1
            End If
        End If
    Next

    '�Ώۃt�@�C���̗�ɐݒ�l���܂܂�Ă��Ȃ������ꍇ�A�������I��
    If FlagWS.Range("B4:B999999").Find(what:="*") Is Nothing Then
        Exit Function
    End If

    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iOpen�j
    Call OpenTextStream

    '�w�b�_�[�쐬
    Call CreateHeader(strFileName)
    
    j = 4
    Do While FlagWS.Cells(j, 2) <> ""
        strTargetSheet = FlagWS.Cells(j, 2)
        strPatternType = FlagWS.Cells(j, 3)
        strYamlKey = SetYamlKey(strTargetSheet, strPatternType)
       
        '�t�@�C�����ݒ�
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
            
    '�����o���֐��Ăяo��
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
    
    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iClose�j
    Call CloseTextStream
    
End Function


'------------------------------'
'�����V�[�g����
'------------------------------'
Public Function JudgeSheet(strFileName As String) As Boolean
    Dim strSerchTarget As String

    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = 99999
        
    '�Ώۃt�@�C���̗�ɐݒ�l���܂߂�Ă��邩����
    strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
    Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
    If Rng Is Nothing Then
        JudgeSheet = False
    Else
        JudgeSheet = True
    End If

End Function


'------------------------------'
'�G�X�P�[�v�����t��
'------------------------------'
Private Function SetEscape(ByVal strTargetVal As String, ByVal strTargetPat As String) As String
    '�ϐ��錾
    Dim EscapeList As Variant
    Dim el As Variant
    
    Select Case strTargetPat
        Case "grep"
            '�G�X�P�[�v�Ώۂ̕�������i�[�i�V�F���j
            EscapeList = Array("[", "]")
        Case "match"
            '�G�X�P�[�v�Ώۂ̕�������i�[�iRuby�j
            EscapeList = Array("$", "*", """")
        Case "block"
            '�G�X�P�[�v�Ώۂ̕�������i�[�i�q�A�h�L�������g�j
            EscapeList = Array("""")
        Case Else
    End Select
    
    '�G�X�P�[�v������t��
    For Each el In EscapeList
        strTargetVal = Replace(strTargetVal, el, "�" + el)
    Next
    
    SetEscape = strTargetVal
End Function


'------------------------------'
'�Z�������s�R�[�h�ϊ�
'------------------------------'
Private Function SetLineBreak(ByVal strTarget As String) As String
      SetLineBreak = Replace(strTarget, vbLf, "�n")
End Function

               
'------------------------------'
'Yaml_Key�ݒ�
'------------------------------'
Private Function SetYamlKey(ByVal strTargetSheet As String, ByVal strPatternType As String) As String
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)

    Select Case strPatternType
        Case "A", "B", "C", "D", "H", "J"
            strYamlKey = WS.Cells(6, 3).Value
            strYamlKey = Mid(strYamlKey, InStrRev(strYamlKey, "/") + 1)
            strYamlKey = Replace(strYamlKey, ".", "_")
            strYamlKey = Replace(strYamlKey, "-", "_")
            strYamlKey = StrConv(strYamlKey, vbNarrow + vbProperCase)

        Case "E"
            If strTargetSheet = "�y�O���[�v�z" Then
                strYamlKey = "Group"
            ElseIf strTargetSheet = "�y���[�U�z" Then
                strYamlKey = "User"
            ElseIf strTargetSheet = "�y�t�@�C���z�z�z" Then
                strYamlKey = "File"
            ElseIf strTargetSheet = "�y�f�B���N�g���쐬�z" Then
                strYamlKey = "Directory"
            ElseIf strTargetSheet = "�y�p�b�P�[�W�z" Then
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
'TargetList�t�@�C���쐬
'------------------------------'
Private Function CreateTargetList()
    Dim i As Long
    
    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iOpen�j
    Call OpenTextStream
    
    '�R�[�h�����o��
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
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set ListWS = ThisWorkbook.Worksheets("�R�[�h�ꗗ")
    
    '�J�n�ʒu�ݒ�
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
            
            '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
            Set WS = ThisWorkbook.Worksheets("���V�s�ꗗ")
            
            '�J�n�s�ݒ�
            i = 4
    
            '�I���s�ݒ�
            EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
            
            'CookBook���擾
            Do While i < EndRow
                If WS.Cells(i, 3).Value <> "" Then
                    strYamlCode = "     - '" + WS.Cells(i, 3).Value + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                End If
                i = i + 1
            Loop
            ActiveColumn = ActiveColumn + 1
        Loop
        '������
        ActiveColumn = 4
        
        ActiveRow = ActiveRow + 1
    Loop
    
    '�����o���֐��Ăяo��
    Call FileOutput("targetList.yml")
    
    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iClose�j
    Call CloseTextStream

End Function

'------------------------------'
'�w�b�_�[�쐬
'------------------------------'
Private Function CreateHeader(ByVal strFileName As String)
    '�R�[�h�����o��
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
'�p�^�[��A�R�[�h����
'------------------------------'
Private Function CreatePatternA(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml�uname�vkey���̗���擾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            '�R�����g�A�E�g������
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
                
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml�uname�v�擾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'�p�^�[��B�R�[�h����
'------------------------------'
Private Function CreatePatternB(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�u���b�N���A�R�����g�A�E�g���̔���
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                '�R�[�h�����o���ilines�j
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���igrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                '�R�[�h�����o���igrep_val�j
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'�p�^�[��C�R�[�h����
'------------------------------'
Private Function CreatePatternC(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml�uname�vkey���̗���擾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '�ύX���ڏ���
    '------------------------------'
    '�ύX���ڏ����J�n�s�ݒ�
    ActiveRow = 6
    
    '�ύX���ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            '�R�����g�A�E�g������
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml�uname�v�擾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '�ǉ����ڏ���
    '------------------------------'
    '�ǉ����ڏ����J�n�s�ݒ�
    ActiveRow = ActiveRow + 3
    
    '�ǉ����ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ǉ����ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�u���b�N���A�R�����g�A�E�g���̔���
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                '�R�[�h�����o���ilines�j
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���igrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                '�R�[�h�����o���igrep_val�j
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'�p�^�[��D�R�[�h�����igrub.conf�j
'------------------------------'
Private Function CreatePatternD(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '�ϐ��錾
    Dim lines() As String
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml�uname�vkey���̗���擾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '�ύX���ڏ���
    '------------------------------'
    '�ύX���ڏ����J�n�s�ݒ�
    ActiveRow = 6
    
    '�ύX���ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            '�R�����g�A�E�g������
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml�uname�v�擾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '�ǉ�����(kernel�s)
    '------------------------------'
    '�ǉ����ڏ����J�n�s�ݒ�
    ActiveRow = ActiveRow + 3
    
    '�ǉ����ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ǉ�����(kernel�s)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�����g�A�E�g���̔���
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: 'kernel'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���igrep_val�j
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '�ǉ�����(�t�@�C������)
    '------------------------------'
    '�ǉ����ڏ����J�n�s�ݒ�
    ActiveRow = ActiveRow + 3
    
    '�ǉ����ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ǉ�����(�t�@�C������)�j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�u���b�N���A�R�����g�A�E�g���̔���
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                '�R�[�h�����o���ilines�j
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���igrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                '�R�[�h�����o���igrep_val�j
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'�p�^�[��E�R�[�h����
'------------------------------'
Private Function CreatePatternE(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    Dim HashArray() As String
    Dim HashCount As Long
    Dim HashColumn As Long
    Dim HashRow As Long
    Dim strHashName As String
    Dim strHashValue As String
    Dim ha As Variant
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Hash���擾
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash���ʒu�擾
    HashRow = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine

    Select Case strTargetSheet
        Case "�y���[�U�z"
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
                    '�R�[�h�����o���i-key�j
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yaml�p�^�[������
                    '�R�����g�A�E�g��������
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    '�R�[�h�����o���ipattern�j
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '������
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash�����s���[�v����
                    Do While Count < HashCount
                        '�R�[�h�����o���iHash���j
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
        Case "�y�p�b�P�[�W�z"
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
                    '�R�[�h�����o���i-key�j
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yaml�p�^�[������
                    '�R�����g�A�E�g��������
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    '�R�[�h�����o���ipattern�j
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '������
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash�����s���[�v����
                    Do While Count < HashCount
                        '�R�[�h�����o���iHash���j
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
            Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
                    '�R�[�h�����o���i-key�j
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yaml�p�^�[������
                    '�R�����g�A�E�g��������
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    '�R�[�h�����o���ipattern�j
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '������
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash��", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash�����s���[�v����
                    Do While Count < HashCount
                        '�R�[�h�����o���iHash���j
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
'�p�^�[��F�R�[�h�����i�����N���T�[�r�X�j
'------------------------------'
Private Function CreatePatternF(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
        If WS.Cells(ActiveRow, 3) <> "" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            If WS.Cells(ActiveRow, ActiveFileColumn) = "" Or WS.Cells(ActiveRow, ActiveFileColumn) = "-" Then
                '�R�[�h�����o���ipattern�j
                strYamlCode = "    :pattern: 'noexist'"
                WriteStream.WriteText strYamlCode, adWriteLine
        
            ElseIf WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                '�R�[�h�����o���ipattern�j
                strYamlCode = "    :pattern: 'exist'"
                WriteStream.WriteText strYamlCode, adWriteLine
            
                Count = 0
                Do While Count < 7
                    '�R�[�h�����o���irunlevel�j
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
'�p�^�[��G�R�[�h�����ixinetd�T�[�r�X�j
'------------------------------'
Private Function CreatePatternG(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "�ȏ�"
        If WS.Cells(ActiveRow, 3) <> "" Then
    
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            '�R�[�h�����o���ipattern�j
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
'�p�^�[��H�R�[�h�����iHosts�j
'------------------------------'
Private Function CreatePatternH(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 4) <> "�ȏ�"
        If WS.Cells(ActiveRow, 4) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            ActiveColumn = 5
            Do While ActiveColumn < 8
                If WS.Cells(ActiveRow, ActiveColumn) <> "" Then
                    '�R�[�h�����o���i-key�j
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
            
                    '�R�[�h�����o���iipaddress�j
                    strYamlCode = "    :ipaddress: '" + WS.Cells(ActiveRow, 4) + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '�R�[�h�����o���iHostname)
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
'�p�^�[��I�R�[�h�����i���[�U���ϐ��j
'------------------------------'
Private Function CreatePatternI(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    '�R�[�h�����o���ikey�j
    If strEvFlag <> strFileName Then
        strYamlCode = strYamlKey + ":"
        WriteStream.WriteText strYamlCode, adWriteLine
        strEvFlag = strFileName
    End If
    
    '�R�[�h�����o���i-key�j
    strYamlCode = "  - " + LCase(strYamlKey) + " :"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    '�R�[�h�����o���ipath�j
    strYamlCode = "    :path: '" + WS.Cells(6, 5) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    '�R�[�h�����o���ifile_name�j
    strYamlCode = "    :file_name: '" + WS.Cells(6, 6) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine

    '�R�[�h�����o���ivalue�j
    strYamlCode = "    :value:"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "��" Then
            '�R�[�h�����o���i-val�j
            strYamlCode = "      - val :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            '�u���b�N���A�R�����g�A�E�g���̔���
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "        :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
                
            If strYamlPattern = "block" Then
                '�R�[�h�����o���ilines�j
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "        :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���igrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "        :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                '�R�[�h�����o���imatch_val�j
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "        :match_val: """ + strYamlValue + """" + "�n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                '�R�[�h�����o���igrep_val�j
                strYamlCode = "        :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine

                '�R�[�h�����o���imatch_val�j
                strYamlCode = "        :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If

        End If
        
        ActiveRow = ActiveRow + 1
    Loop
End Function


'------------------------------'
'�p�^�[��J�R�[�h�����isysctl.conf�j
'------------------------------'
Private Function CreatePatternJ(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�R�[�h�����o���ikey�j
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yaml�uname�vkey���̗���擾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '�ύX���ڏ���
    '------------------------------'
    '�ύX���ڏ����J�n�s�ݒ�
    ActiveRow = 6
    
    '�ύX���ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ύX���ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            '�R�����g�A�E�g������
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yaml�uname�v�擾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '�ǉ����ڏ���
    '------------------------------'
    '�ǉ����ڏ����J�n�s�ݒ�
    ActiveRow = ActiveRow + 3
    
    '�ǉ����ڏ����I���s�ݒ�
    EndRow = WS.Cells.Find("�ȏ�i�ǉ����ځj", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
            '�R�[�h�����o���i-key�j
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�����g�A�E�g���̔���
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            '�R�[�h�����o���ipattern�j
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���iname�j
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���igrep_val�j
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            '�R�[�h�����o���imatch_val�j
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function


Sub test()
    strTargetSheet = "�y�N���X�N���v�g�zrc.local"
    strFileName = "stprdb01"
    strYamlKey = "Rc_local"
    '�t�H���_����
    OutputPath = ThisWorkbook.Worksheets("���C��").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "�o�͐�t�H���_�Ɍ�肪����܂��B", vbCritical
        Exit Sub
    End If

    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iOpen�j
    Call OpenTextStream
    Call CreatePatternC(strTargetSheet, strFileName, strYamlKey)
        
    '�����o���֐��Ăяo��
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
        
    '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iClose�j
    Call CloseTextStream

End Sub

