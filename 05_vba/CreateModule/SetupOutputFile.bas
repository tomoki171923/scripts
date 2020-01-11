Attribute VB_Name = "SetupOutputfile"
Public strFileName As String
Public WriteStream As Object

'------------------------------'
'�e�L�X�g�X�g���[���̃I�u�W�F�N�g�ݒ�iOpen�j
'------------------------------'
Public Function OpenTextStream()
    ' �e�L�X�g�X�g���[���̃I�u�W�F�N�g�쐬 & �ݒ�(�G���R�[�h�FUTF8�A���s�R�[�h�FLF)
    Set WriteStream = CreateObject("ADODB.Stream")
    With WriteStream
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
    End With

    ' �e�L�X�g�X�g���[���̃I�u�W�F�N�g���J��
    WriteStream.Open
End Function

'------------------------------'
'�e�L�X�g�X�g���[���̃I�u�W�F�N�g�ݒ�iClose�j
'------------------------------'
Public Function CloseTextStream()
    ' �e�L�X�g�X�g���[���̃I�u�W�F�N�g�����
    WriteStream.Close
    ' �I�u�W�F�N�g�N���A
    Set WriteStream = Nothing
End Function

'------------------------------'
'�t�@�C�������o��
'------------------------------'
Public Function FileOutput(ByVal strFileName As String)
    Dim objFSO As Object
    Dim Result As Long
    Dim strFullPath As String
    
    '�����Ŏ󂯎�����t�@�C�����Ńe�L�X�g�X�g���[�������o��
    'PATH�w��
    strFullPath = ThisWorkbook.Worksheets("���C��").Range("C5").Value + strFileName
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' �t�@�C�����݊m�F
    If objFSO.FileExists(strFullPath) = True Then
        Result = MsgBox("�w�肵���t�H���_�ɓ�����" + strFileName + "�����݂��܂��B�㏑�����܂����B", vbYesNo)
        If Result = 6 Then
            ' �R���t�B�O�t�@�C�������o��
            WriteStream.SaveToFile (strFullPath), adSaveCreateOverWrite
            MsgBox strFileName + "���㏑�����܂����B"
        Else
            MsgBox strFileName + "�̏㏑���𒆎~���܂����B"
            Set objFSO = Nothing
        End If
    Else
        ' �R���t�B�O�t�@�C�������o��
        WriteStream.SaveToFile (strFullPath), adSaveCreateOverWrite
        MsgBox strFileName + "���쐬���܂����B"
    End If
    
    ' �I�u�W�F�N�g�N���A
    Set objFSO = Nothing

End Function

