Attribute VB_Name = "SetupFolder"
Public OutputPath As String

'------------------------------'
' ƒtƒHƒ‹ƒ_QÆƒ_ƒCƒAƒƒO
'------------------------------'
Sub BrowesFolder()
    Dim SerchChell As Range
    Set SerchChell = ThisWorkbook.Worksheets("ƒƒCƒ“").Cells.Find("o—ÍêŠ", LookIn:=xlValues, LookAt:=xlWhole)
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            SerchChell.Offset(0, 1) = .SelectedItems(1) + "€"
        End If
    End With
End Sub

'------------------------------'
' QÆƒtƒHƒ‹ƒ_”»’è
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
