Attribute VB_Name = "SetupFolder"
Public OutputPath As String

'------------------------------'
' フォルダ参照ダイアログ
'------------------------------'
Sub BrowesFolder()
    Dim SerchChell As Range
    Set SerchChell = ThisWorkbook.Worksheets("メイン").Cells.Find("出力場所", LookIn:=xlValues, LookAt:=xlWhole)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            SerchChell.Offset(0, 1) = .SelectedItems(1) + "�"
        End If
    End With
End Sub

'------------------------------'
' 参照フォルダ判定
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
