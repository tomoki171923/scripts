Attribute VB_Name = "CreateErb"
Dim TargetSheet As Worksheet
Dim strPatternType As String


Sub ControllerErb()
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
    StartColumn = 3
    StartRow = 5
    
    Do While WS.Cells(StartRow, StartColumn).Value <> ""
        strTargetName = WS.Cells(StartRow, StartColumn).Value
        Call Tyuukann(strTargetName)
            
        StartColumn = StartColumn + 1
        Do While WS.Cells(StartRow, StartColumn).Value <> ""
            strTargetName = WS.Cells(StartRow, StartColumn).Value
            Call Tyuukann(strTargetName)
            StartColumn = StartColumn + 1
        Loop
        '������
        StartColumn = 3
        strTargetName = ""
        StartRow = StartRow + 1
    Loop
End Sub
    
Private Function Tyuukann(ByVal strTargetName As String)
    '���[�N�V �[�g�I�u�W�F�N�g�ݒ聕�N���A
    Set FlagWS = ThisWorkbook.Worksheets("�����V�[�g")
    FlagWS.Range("B4:D999999").ClearContents

    i = 4
    '�����V�[�g�ɏ����o��
    For Each TargetSheet In Worksheets
        If InStr(TargetSheet.Name, "hosts") > 0 Then
            If CheckSheet(strTargetName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = "hosts"
                i = i + 1
            End If
        ElseIf InStr(TargetSheet.Name, "���[�U���ϐ��i�ʁj") > 0 Then
            If CheckSheet(strTargetName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = "���[�U���ϐ�"
                i = i + 1
            End If
        End If
    Next

    '�Ώۃt�@�C���̗�ɐݒ�l���܂܂�Ă��Ȃ������ꍇ�A�������I��
    If FlagWS.Range("B4:B999999").Find(what:="*") Is Nothing Then
        Exit Function
    End If

    j = 4
    Do While FlagWS.Cells(j, 2) <> ""
        '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iOpen�j
        Call OpenTextStream

        strTargetSheet = FlagWS.Cells(j, 2)
        strPatternType = FlagWS.Cells(j, 3)
        '�t�@�C�����ݒ�
        Select Case strPatternType
            Case "hosts"
                Call CreateHosts(strTargetSheet, strTargetName)
            Case "���[�U���ϐ�"
                Call CreateEnviromentVal(strTargetSheet, strTargetName)
            Case Else
        End Select
            
        ' �����o���֐��Ăяo��
        strFileName = strFileName + ".erb"
        Call FileOutput(strFileName)
        
        '�e�L�X�g�X�g���[���̃I�u�W�F�N�g��ݒ肷��iClose�j
        Call CloseTextStream
        
        j = j + 1
    Loop
    
End Function


Function CreateHosts(ByVal strTargetSheet As String, ByVal strTargetName As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�w�b�_�[�����o��
    WriteStream.WriteText "127.0.0.1   localhost localhost.localdomain localhost4 localhost4.localdomain4", adWriteLine
    WriteStream.WriteText "#::1         localhost localhost.localdomain localhost6 localhost6.localdomain6", adWriteLine
    WriteStream.WriteText "", adWriteLine
    WriteStream.WriteText "", adWriteLine
            
    '�ǉ����ڗ񃋁[�v����
    Do While WS.Cells(ActiveRow, 4) <> "�ȏ�"
        If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
            strChefCode = WS.Cells(ActiveRow, 4).Value + vbTab + WS.Cells(ActiveRow, 5).Value + vbTab + WS.Cells(ActiveRow, 6).Value + vbTab + WS.Cells(ActiveRow, 7).Value + vbTab + WS.Cells(ActiveRow, 8).Value
            WriteStream.WriteText strChefCode, adWriteLine
  
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '�o�̓t�@�C�����ݒ�
    strFileName = strTargetName + "_hosts"
    
End Function




Function CreateEnviromentVal(ByVal strTargetSheet As String, ByVal strTargetName As String)
    '���[�N�V�[�g�I�u�W�F�N�g�ݒ�
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = WS.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�w�b�_�[�����o��
    Select Case WS.Cells(6, 6)
        Case ".bash_profile"
            WriteStream.WriteText "# .bash_profile", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# Get the aliases and functions", adWriteLine
            WriteStream.WriteText "if [ -f ~/.bashrc ]; then", adWriteLine
            WriteStream.WriteText "# .bash_profile", adWriteLine
            WriteStream.WriteText "    . ~/.bashrc", adWriteLine
            WriteStream.WriteText "fi", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific environment and startup programs", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "PATH=$PATH:$HOME/bin", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "export PATH", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".bashrc"
            WriteStream.WriteText "# .bashrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# Source global definitions", adWriteLine
            WriteStream.WriteText "if [ -f /etc/bashrc ]; then", adWriteLine
            WriteStream.WriteText "    . /etc/bashrc", adWriteLine
            WriteStream.WriteText "fi", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".tcshrc"
            WriteStream.WriteText "# .tcshrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "set prompt='[%n@%m %c]# '", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case ".cshrc"
            WriteStream.WriteText "# .cshrc", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "# User specific aliases and functions", adWriteLine
            WriteStream.WriteText "", adWriteLine
            WriteStream.WriteText "alias rm='rm -i'", adWriteLine
            WriteStream.WriteText "alias cp 'cp -i'", adWriteLine
            WriteStream.WriteText "alias mv 'mv -i'", adWriteLine
            WriteStream.WriteText "", adWriteLine
        Case Else
    End Select
    
    'root���菈��
    If WS.Cells(6, 3) = "root" Then
        WriteStream.WriteText "export PATH=chefdk/frhu", adWriteLine
        WriteStream.WriteText "", adWriteLine
    End If
    
    
    '�ǉ����ڗ񃋁[�v����
    Do While WS.Cells(ActiveRow, 12) <> "�ȏ�"
        If WS.Cells(ActiveRow, ActiveFileColumn).Value = "��" Then
            strChefCode = WS.Cells(ActiveRow, 12).Value
            WriteStream.WriteText strChefCode, adWriteLine
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '�o�̓t�@�C�����ݒ�
    strFileName = strTargetName + "_" + WS.Cells(6, 3).Value + WS.Cells(6, 6).Value

End Function

'------------------------------'
'�����V�[�g����
'------------------------------'
Public Function CheckSheet(strTargetName As String) As Boolean

    '�Ώۃt�@�C���̗��ݒ�
    ActiveFileColumn = TargetSheet.Cells.Find(strTargetName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '�J�n�s�ݒ�
    ActiveRow = 6
    
    '�I���s�ݒ�
    EndRow = TargetSheet.Cells.Find("�ȏ�", LookIn:=xlValues, LookAt:=xlWhole).Row
        
    '�Ώۃt�@�C���̗�ɐݒ�l���܂߂�Ă��邩����
    strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
    Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="��")
    If Rng Is Nothing Then
        CheckSheet = False
    Else
        CheckSheet = True
    End If

End Function
