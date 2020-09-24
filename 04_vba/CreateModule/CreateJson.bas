Attribute VB_Name = "CreateJson"
'Option Explicit
    '•Ï”éŒ¾iƒOƒ[ƒoƒ‹j
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
    
    '’è”éŒ¾iƒOƒ[ƒoƒ‹j
    Public Const Dquote As String = """"
    Public Const indent1 As String = "  "
    Public Const indent2 As String = "    "
    Public Const indent3 As String = "      "
    Public Const indent4 As String = "        "
    Public Const indent5 As String = "          "
    
'------------------------------'
'ƒRƒ“ƒgƒ[ƒ‰[
'------------------------------'
Sub ControllerJson()
    '•Ï”éŒ¾
    Dim StartColumn As Long
    Dim StartRow As Long
    
    'ƒtƒHƒ‹ƒ_”»’è
    OutputPath = ThisWorkbook.Worksheets("ƒƒCƒ“").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "o—ÍæƒtƒHƒ‹ƒ_‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbCritical
        Exit Sub
    End If

    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets("ƒR[ƒhˆê——")

    'ŠJnˆÊ’uİ’è
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
        '‰Šú‰»
        StartColumn = 3
        strNodeName = ""
        
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'ƒpƒ^[ƒ“”»’è
'------------------------------'
Private Function JudgePattern(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiOpenj
    Call OpenTextStream
    
    'ƒwƒbƒ_[‚ğì¬‚·‚é
    Call CreateHeader(strFileType, strRoleName, strNodeName)
    
    'ƒtƒ@ƒCƒ‹–¼İ’è
    Select Case strFileType
        Case "Role"
            strFileName = strRoleName
        Case "Node"
            strFileName = strNodeName
        Case Else
    End Select
       
    'ƒ[ƒNƒV [ƒgƒIƒuƒWƒFƒNƒgİ’è•ƒNƒŠƒA
    Set FlagWS = ThisWorkbook.Worksheets("ˆ—ƒV[ƒg")
    FlagWS.Range("B4:D999999").ClearContents
       
    'ˆ—ƒV[ƒgİ’è
    Call JudgeSheets(strFileName)

    'ÅIˆ—ƒV[ƒg–¼æ“¾
    strLastSheet = FlagWS.Range("B3").End(xlDown).Value
        
    a = 4
    Do While FlagWS.Cells(a, 2) <> ""
        strTargetSheet = FlagWS.Cells(a, 2)
        strPatternType = FlagWS.Cells(a, 3)
    
        'ƒpƒ^[ƒ“”»’è
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
    
    'ƒtƒbƒ^[‚ğì¬‚·‚é
    Call CreateFooter(strFileType)

    ' ‘‚«o‚µŠÖ”ŒÄ‚Ño‚µ
    strFileName = strFileName + ".json"
    Call FileOutput(strFileName)

    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiClosej
    Call CloseTextStream

End Function

'------------------------------'
'ˆ—ƒV[ƒg”»’è
'------------------------------'
Private Function JudgeSheets(ByVal strFileName As String) As String
    '•Ï”éŒ¾
    Dim PatternCell As Range
    i = 4
    'ƒV[ƒgƒ‹[ƒv
    For Each TargetSheet In ThisWorkbook.Worksheets
        'ƒpƒ^[ƒ“–¼ŒŸõ
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        If PatternCell Is Nothing Then
            strPatternType = ""
        Else
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
        End If
    
        'ŠJnsİ’è
        ActiveRow = 6
 
        '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚Éİ’è’l‚ªŠÜ‚ß‚ê‚Ä‚¢‚é‚©”»’è
        Select Case strPatternType
            Case "A"
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
                
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
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
        
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="›")
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
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈãi’Ç‰Á€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚É"›"‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚©”»’è
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="›")

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
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row
            
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
                If Rng Is Nothing Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈãi’Ç‰Á€–Ú(kernels)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚É"›"‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚©”»’è
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="›")

                If Rng Is Nothing And strFlag = "off" Then
                    strFlag = "off"
                Else
                    strFlag = "on"
                End If
                
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈãi’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚É"›"‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚©”»’è
                strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
                Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="›")

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
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row

                'Template‚Ì—ñ‚ğİ’è
                TemplateColumn = 7
    
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚Éİ’è’l‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚©”»’è
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
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
                ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
                'I—¹sİ’è
                EndRow = TargetSheet.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
    
                '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚Éİ’è’l‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚©”»’è
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
'ƒGƒXƒP[ƒv•¶š•t‰Á
'------------------------------'
Private Function AddEscape(ByVal strTarget As String) As String
    'ƒ_ƒuƒ‹ƒNƒH[ƒe[ƒVƒ‡ƒ“‚ª‚ ‚éê‡ƒGƒXƒP[ƒv•¶š’Ç‰Á
    AddEscape = Replace(strTarget, """", "€""")
End Function


'------------------------------'
'ƒwƒbƒ_[ì¬
'------------------------------'
Private Function CreateHeader(ByVal strFileType As String, ByVal strRoleName As String, ByVal strNodeName As String)
        'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets("ƒŒƒVƒsˆê——")
    
    'ŠJnsİ’è
    ActiveRow = 4
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
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
        
        
            'CookBook–¼æ“¾
            Do While ActiveRow < EndRow
                If WS.Cells(ActiveRow, 3).Value <> "" Then
                    If WS.Cells(ActiveRow + 1, 3).Value <> "" Then
                        'ˆ—s”İ’è
                        Line = 1
                    Else
                        'ˆ—s”İ’è
                        LineCount = WS.Cells(ActiveRow, 3).End(xlDown).Row
                        Line = LineCount - ActiveRow
                    End If
                    strCookbookName = WS.Cells(ActiveRow, 3).Value
                
                    j = 1
                    'Recipe–¼æ“¾
                    Do While j <= Line
                        strRecipeName = WS.Cells(ActiveRow, 4).Value
                        
                        If strRecipeName <> "" Then
                            'ƒJƒ“ƒ}”»’èˆ—irun_list“àj
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
        'Nodeƒtƒ@ƒCƒ‹ˆ—
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
            
            'CookBook–¼æ“¾
            Do While WS.Cells(ActiveRow, 3).Value <> "ˆÈã"
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
'ƒtƒbƒ^[ì¬
'------------------------------'
Private Function CreateFooter(ByVal strFileType As String)
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            WriteStream.WriteText indent2
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            WriteStream.WriteText indent1
            WriteStream.WriteText "}", adWriteLine
            WriteStream.WriteText "}", adWriteLine
        Case Else
    End Select
End Function


'------------------------------'
'ƒpƒ^[ƒ“AƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternA(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute–¼‚Ì—ñ‚ğæ“¾
    AttributeColumn = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                   
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
'ƒpƒ^[ƒ“BƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternB(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute–¼æ“¾
    AttributeRow = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            'Attribute–¼‘‚«o‚µ
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            'Attribute–¼‘‚«o‚µ
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“CƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternC(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("ˆÈãi’Ç‰Á€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute–¼‚Ì—ñ‚ğæ“¾
    AttributeColumn = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi•ÏX€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
            '’Ç‰Á€–Úˆ—
            '------------------------------'
            '’Ç‰Á€–Úˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="›")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute–¼‘‚«o‚µ
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
                            
        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi•ÏX€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
            '’Ç‰Á€–Úˆ—
            '------------------------------'
            '’Ç‰Á€–Úˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng = WS.Range(strSerchTarget).Find(what:="›")
            If Rng Is Nothing Then
                Exit Function
            End If
            'Attribute–¼‘‚«o‚µ
            strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“DƒR[ƒh¶¬igrub.confj
'------------------------------'
Private Function CreatePatternD(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow2 = WS.Cells.Find("ˆÈãi’Ç‰Á€–Ú(kernels)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    EndRow3 = WS.Cells.Find("ˆÈãi’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'Attribute–¼‚Ì—ñ‚ğæ“¾
    AttributeColumn = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
       
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "ˆÈãi•ÏX€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
            '’Ç‰Á€–Ú(kernels)ˆ—
            '------------------------------'
            '’Ç‰Á€–Ú(kernels)ˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
            If Not Rng2 Is Nothing Then
                'Attribute–¼‘‚«o‚µ
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '’Ç‰Á€–Ú(kernels)—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Ú(kernels)j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)ˆ—
            '------------------------------'
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)ˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
            If Not Rng3 Is Nothing Then
                'Attribute–¼‘‚«o‚µ
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent4 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If

        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            '------------------------------'
            '•ÏX€–Úˆ—
            '------------------------------'
            Do While WS.Cells(ActiveRow, 12).Value <> "ˆÈãi•ÏX€–Új"
                If WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                    'Attribute–¼Eİ’è’læ“¾A‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                    strAttributeName = WS.Cells(ActiveRow, AttributeColumn).Value
                    strAttributeValue = AddEscape(WS.Cells(ActiveRow, ActiveFileColumn).Value)
                    
                    strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strAttributeValue + Dquote
                    WriteStream.WriteText strChefCode
           
                    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="*")
                    strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
                    Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
                    'ƒAƒNƒeƒBƒuƒV[ƒg‚ªÅIˆ—ƒV[ƒg‚©‚ÂƒAƒNƒeƒBƒuƒZƒ‹ˆÈ‰º‚É’l‚ª‚È‚¢ê‡
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
            '’Ç‰Á€–Ú(kernels)ˆ—
            '------------------------------'
            '’Ç‰Á€–Ú(kernels)ˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
            Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
            If Not Rng2 Is Nothing Then
                'Attribute–¼‘‚«o‚µ
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '’Ç‰Á€–Ú(kernels)—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Ú(kernels)j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    'ƒR[ƒhì¬•ƒGƒXƒP[ƒv•¶š•t‰Á
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow2, ActiveFileColumn).Address
                    Set Rng2 = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng2 Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
            If strTargetSheet = strLastSheet And Rng3 Is Nothing Then
                strChefCode = indent2 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent2 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
            
            '------------------------------'
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)ˆ—
            '------------------------------'
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)ˆ—ŠJnsİ’è
            ActiveRow = ActiveRow + 2
            
            'Attribute‘‚«o‚µ”»’è
            strSerchTarget = WS.Cells(ActiveRow, ActiveFileColumn).Address + ":" + WS.Cells(EndRow3, ActiveFileColumn).Address
            Set Rng3 = WS.Range(strSerchTarget).Find(what:="›")
            If Not Rng3 Is Nothing Then
                'Attribute–¼‘‚«o‚µ
                strAttributeName = WS.Cells(ActiveRow - 1, AttributeColumn).Value
                strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
                WriteStream.WriteText strChefCode, adWriteLine
            Else
                Exit Function
            End If
            
            '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 12) <> "ˆÈãi’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)j"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    strChefCode = indent3 + Dquote + AddEscape(WS.Cells(ActiveRow, 12).Value) + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "", adWriteLine
                    Else
                        WriteStream.WriteText ",", adWriteLine
                    End If
                End If
                ActiveRow = ActiveRow + 1
            Loop
            
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“EƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternE(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '•Ï”éŒ¾
    Dim HashRow As Long
    Dim HashColumn As Long
    Dim HashCount As Long
    Dim strHashValue As String
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Attribute–¼æ“¾
    AttributeRow = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    AttributeColumn = WS.Cells.Find("Attribute–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
    strAttributeName = WS.Cells(AttributeRow, AttributeColumn).Value
    
    'HashŒÂ”æ“¾
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash–¼ˆÊ’uæ“¾
    HashRow = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            'Attribute–¼‘‚«o‚µ
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    strChefCode = indent4 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '‰Šú‰»
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash”•ªsƒ‹[ƒvˆ—
                    Do While Count < HashCount
                        'ƒnƒbƒVƒ…ƒL[AƒnƒbƒVƒ…’læ“¾@‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                        strHashValue = WS.Cells(ActiveRow, ActiveColumn).Value
                        strHashValue = AddEscape(strHashValue)
    'š
                        strChefCode = indent5 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        'ƒJƒ“ƒ}”»’èˆ—iHash“àj
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
                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
            If strTargetSheet = strLastSheet Then
                strChefCode = indent3 + "]"
                WriteStream.WriteText strChefCode, adWriteLine
                Exit Function
            Else
                strChefCode = indent3 + "],"
                WriteStream.WriteText strChefCode, adWriteLine
            End If
        
        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            'Attribute–¼‘‚«o‚µ
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": ["
            WriteStream.WriteText strChefCode, adWriteLine
            
            '’Ç‰Á€–Ú—ñƒ‹[ƒvˆ—
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "›" Then
                    strChefCode = indent3 + "{"
                    WriteStream.WriteText strChefCode, adWriteLine
                    '‰Šú‰»
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash”•ªsƒ‹[ƒvˆ—
                    Do While Count < HashCount
                        'ƒnƒbƒVƒ…ƒL[AƒnƒbƒVƒ…’læ“¾@‘‚«o‚µ&ƒGƒXƒP[ƒv•¶š•t‰Á
                        strHashValue = AddEscape(WS.Cells(ActiveRow, ActiveColumn).Value)
                        strChefCode = indent4 + Dquote + WS.Cells(HashRow, HashColumn).Value + Dquote + ": " + Dquote + strHashValue + Dquote
                        WriteStream.WriteText strChefCode
                        
                        'ƒJƒ“ƒ}”»’èˆ—iHash“àj
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
                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
                    strSerchTarget = WS.Cells(ActiveRow + 1, ActiveFileColumn).Address + ":" + WS.Cells(EndRow, ActiveFileColumn).Address
                    Set Rng = WS.Range(strSerchTarget).Find(what:="›")
                    If Rng Is Nothing Then
                        WriteStream.WriteText "}", adWriteLine
                    Else
                        WriteStream.WriteText "},", adWriteLine
                    End If
                    
                End If
                ActiveRow = ActiveRow + 1
            Loop
            'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“FƒR[ƒh¶¬i©“®‹N“®ƒT[ƒrƒXj
'------------------------------'
Private Function CreatePatternF(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '•Ï”éŒ¾
    Dim strRunLevel As String
    Dim TemplateColumn As Long
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Template‚Ì—ñ‚ğİ’è
    TemplateColumn = 7
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
        
            '------------------------------'
            'Service AddƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent3 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'ƒJƒ“ƒ}”»’èˆ—iHash“àj
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "ˆÈã"
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
            'Service DeleteƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent3 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'ƒJƒ“ƒ}”»’èˆ—iHash“àj
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "ˆÈã"
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
            'Service Change(on)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                'Attribute‘‚«o‚µ”»’è
                If AttWrite = 0 Then
                    'Attribute–¼‘‚«o‚µ
                    strChefCode = indent3 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ƒ‰ƒ“ƒŒƒxƒ‹‚ÌŒÂ”•ªƒ‹[ƒvˆ—
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

                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
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
            'Service Change(off)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                'Attribute‘‚«o‚µ”»’è
                If AttWrite = 0 Then
                    'Attribute–¼‘‚«o‚µ
                    strChefCode = indent3 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ƒ‰ƒ“ƒŒƒxƒ‹‚ÌŒÂ”•ªƒ‹[ƒvˆ—
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

                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
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
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
        
            '------------------------------'
            'Service AddƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, TemplateColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, ActiveFileColumn).Value, 3) = "on")) Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent2 + Dquote + "att_sv_add_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'ƒJƒ“ƒ}”»’èˆ—iHash“àj
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "ˆÈã"
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
            'Service DeleteƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "-" And ((Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "off" Or Mid(WS.Cells(ActiveRow, TemplateColumn).Value, 3) = "on")) Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent2 + Dquote + "att_sv_del_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                
                    'ƒJƒ“ƒ}”»’èˆ—iHash“àj
                    Line = ActiveRow + 1
                    HashWrite = 0
                    Do While WS.Cells(Line, 3).Value <> "ˆÈã"
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
            'Service Change(on)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                'Attribute‘‚«o‚µ”»’è
                If AttWrite = 0 Then
                    'Attribute–¼‘‚«o‚µ
                    strChefCode = indent2 + Dquote + "att_sv_chg_on_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ƒ‰ƒ“ƒŒƒxƒ‹‚ÌŒÂ”•ªƒ‹[ƒvˆ—
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

                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
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
            'Service Change(off)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
  
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                'Attribute‘‚«o‚µ”»’è
                If AttWrite = 0 Then
                    'Attribute–¼‘‚«o‚µ
                    strChefCode = indent2 + Dquote + "att_sv_chg_off_service" + Dquote + ": ["
                    WriteStream.WriteText strChefCode, adWriteLine
                    AttWrite = 1
                End If

                'ƒ‰ƒ“ƒŒƒxƒ‹‚ÌŒÂ”•ªƒ‹[ƒvˆ—
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

                    'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
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
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“GƒR[ƒh¶¬ixinetdƒT[ƒrƒXj
'------------------------------'
Private Function CreatePatternG(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            '------------------------------'
            'Xinetd Service Change(on)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent3 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
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
            'Xinetd Service Change(off)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent3 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent4 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
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
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            '------------------------------'
            'Xinetd Service Change(on)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0

            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "on" Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent2 + Dquote + "att_sv_chg_on_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
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
            'Xinetd Service Change(off)ƒR[ƒhì¬
            '------------------------------'
            ActiveRow = 6
            AttWrite = 0
 
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn).Value = "off" Then
                    'Attribute‘‚«o‚µ”»’è
                    If AttWrite = 0 Then
                        'Attribute–¼‘‚«o‚µ
                        strChefCode = indent2 + Dquote + "att_sv_chg_off_xinetd_service" + Dquote + ": ["
                        WriteStream.WriteText strChefCode, adWriteLine
                        AttWrite = 1
                    End If

                    strChefCode = indent3 + Dquote + WS.Cells(ActiveRow, 3).Value + Dquote
                    WriteStream.WriteText strChefCode
                    
                    'ƒJƒ“ƒ}”»’èˆ—i”z—ñ“àj
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
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
'ƒpƒ^[ƒ“HƒR[ƒh¶¬ihostsj
'------------------------------'
Private Function CreatePatternH(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    'Attribute–¼İ’è
    strAttributeName = "att_ho_hosts_chg_server"
    
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            strChefCode = indent3 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        '------------------------------'
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            strChefCode = indent2 + Dquote + strAttributeName + Dquote + ": " + Dquote + strFileName + Dquote
            WriteStream.WriteText strChefCode
        Case Else
    End Select
        
    'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
    If strTargetSheet = strLastSheet Then
        WriteStream.WriteText "", adWriteLine
        Exit Function
    Else
        WriteStream.WriteText ",", adWriteLine
    End If

End Function





'------------------------------'
'ƒpƒ^[ƒ“IƒR[ƒh¶¬iƒ†[ƒUŠÂ‹«•Ï”j
'------------------------------'
Private Function CreatePatternI(ByVal strFileType As String, ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strLastSheet As String)
    '•Ï”éŒ¾
    Dim strLastEvSheet As String
    Dim strFirstEvSheet As String
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
    Set FlagWS = ThisWorkbook.Worksheets("ˆ—ƒV[ƒg")
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'Attribute–¼İ’è
    strAttributeName = "att_ev_chg_environment_variable"
    
    'ŠJnEI—¹‚Ìƒ†[ƒUŠÂ‹«•Ï”ƒV[ƒg–¼‚ğæ“¾
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
  
    'ƒtƒ@ƒCƒ‹í—Ş”»’è
    Select Case strFileType
    
        '------------------------------'
        'Roleƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Role"
            'Attribute‘‚«o‚µ”»’è
            If strTargetSheet = strFirstEvSheet Then
                'Attribute–¼‘‚«o‚µ
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
            
            'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent4
                WriteStream.WriteText "}", adWriteLine
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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
        'Nodeƒtƒ@ƒCƒ‹ˆ—
        '------------------------------'
        Case "Node"
            'Attribute‘‚«o‚µ”»’è
            If strTargetSheet = strFirstEvSheet Then
                'Attribute–¼‘‚«o‚µ
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
            
            'ƒJƒ“ƒ}”»’èˆ—iHashŠOj
            If strTargetSheet = strLastEvSheet Then
                WriteStream.WriteText indent3
                WriteStream.WriteText "}", adWriteLine
                'ƒJƒ“ƒ}”»’èˆ—iÅIˆ—j
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



