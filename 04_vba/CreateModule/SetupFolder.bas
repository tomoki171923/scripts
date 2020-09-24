Attribute VB_Name = "SetupFolder"
Public OutputPath As String

'------------------------------'
' �t�H���_�Q�ƃ_�C�A���O
'------------------------------'
Sub BrowesFolder()
    Dim SerchChell As Range
    Set SerchChell = ThisWorkbook.Worksheets("���C��").Cells.Find("�o�͏ꏊ", LookIn:=xlValues, LookAt:=xlWhole)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            SerchChell.Offset(0, 1) = .SelectedItems(1) + "�"
        End If
    End With
End Sub

'------------------------------'
' �Q�ƃt�H���_����
'------------------------------'
Public Function CheckPath(val As String) As Boolean
    Dim path As String
    
    If Trim(val) = "" Then
        CheckPath = False
        Exit Function
    End If
    
    path = Dir(val, vbDirectory)
    
    If Trim(path = "") Then
        CheckPath = False
        Exit Function
    End If
    CheckPath = True
End Function
