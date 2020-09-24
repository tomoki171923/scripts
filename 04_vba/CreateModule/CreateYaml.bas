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
'ƒRƒ“ƒgƒ[ƒ‰
'------------------------------'
Sub ControllerYaml()
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
    Set ListWS = ThisWorkbook.Worksheets("ƒR[ƒhˆê——")

    'TargetListƒtƒ@ƒCƒ‹ì¬
    Call CreateTargetList

    'ŠJnˆÊ’uİ’è
    StartColumn = 3
    StartRow = 5
    
    'Šeyamlƒtƒ@ƒCƒ‹ì¬
    Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
        strFileName = ListWS.Cells(StartRow, StartColumn).Value
        Call JudgePattern(strFileName)
            
        StartColumn = StartColumn + 1
        Do While ListWS.Cells(StartRow, StartColumn).Value <> ""
            strFileName = ListWS.Cells(StartRow, StartColumn).Value
            Call JudgePattern(strFileName)
            StartColumn = StartColumn + 1
        Loop
        '‰Šú‰»
        StartColumn = 3
        strFileName = ""
        StartRow = StartRow + 1
    Loop
    
End Sub

'------------------------------'
'ƒpƒ^[ƒ“”»’è
'------------------------------'
Private Function JudgePattern(ByVal strFileName As String)
    Dim i As Long
    Dim j As Long
    Dim PatternCell As Object
    
    'ƒ[ƒNƒV [ƒgƒIƒuƒWƒFƒNƒgİ’è•ƒNƒŠƒA
    Set FlagWS = ThisWorkbook.Worksheets("ˆ—ƒV[ƒg")
    FlagWS.Range("B4:D999999").ClearContents

    i = 4
    'ˆ—ƒV[ƒg‚É‘‚«o‚µ
    For Each TargetSheet In ThisWorkbook.Worksheets
        'ƒpƒ^[ƒ“–¼ŒŸõ
        Set PatternCell = TargetSheet.Cells.Find("Pattern", LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not PatternCell Is Nothing Then
            strPatternType = TargetSheet.Cells(PatternCell.Row + 1, PatternCell.Column).Value
            
            '‘ÎÛƒtƒ@ƒCƒ‹—ñ‚É•¶š—ñ‚ªŠÜ‚ñ‚Å‚¢‚é‚©”»’è
            If JudgeSheet(strFileName) = True Then
                FlagWS.Cells(i, 2).Value = TargetSheet.Name
                FlagWS.Cells(i, 3).Value = strPatternType
                i = i + 1
            End If
        End If
    Next

    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚Éİ’è’l‚ªŠÜ‚Ü‚ê‚Ä‚¢‚È‚©‚Á‚½ê‡Aˆ—‚ğI—¹
    If FlagWS.Range("B4:B999999").Find(what:="*") Is Nothing Then
        Exit Function
    End If

    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiOpenj
    Call OpenTextStream

    'ƒwƒbƒ_[ì¬
    Call CreateHeader(strFileName)
    
    j = 4
    Do While FlagWS.Cells(j, 2) <> ""
        strTargetSheet = FlagWS.Cells(j, 2)
        strPatternType = FlagWS.Cells(j, 3)
        strYamlKey = SetYamlKey(strTargetSheet, strPatternType)
       
        'ƒtƒ@ƒCƒ‹–¼İ’è
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
            
    '‘‚«o‚µŠÖ”ŒÄ‚Ño‚µ
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
    
    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiClosej
    Call CloseTextStream
    
End Function


'------------------------------'
'ˆ—ƒV[ƒg”»’è
'------------------------------'
Public Function JudgeSheet(strFileName As String) As Boolean
    Dim strSerchTarget As String

    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = TargetSheet.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = 99999
        
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚Éİ’è’l‚ªŠÜ‚ß‚ê‚Ä‚¢‚é‚©”»’è
    strSerchTarget = TargetSheet.Cells(ActiveRow, ActiveFileColumn).Address + ":" + TargetSheet.Cells(EndRow, ActiveFileColumn).Address
    Set Rng = TargetSheet.Range(strSerchTarget).Find(what:="*")
    If Rng Is Nothing Then
        JudgeSheet = False
    Else
        JudgeSheet = True
    End If

End Function


'------------------------------'
'ƒGƒXƒP[ƒv•¶š•t‰Á
'------------------------------'
Private Function SetEscape(ByVal strTargetVal As String, ByVal strTargetPat As String) As String
    '•Ï”éŒ¾
    Dim EscapeList As Variant
    Dim el As Variant
    
    Select Case strTargetPat
        Case "grep"
            'ƒGƒXƒP[ƒv‘ÎÛ‚Ì•¶š—ñ‚ğŠi”[iƒVƒFƒ‹j
            EscapeList = Array("[", "]")
        Case "match"
            'ƒGƒXƒP[ƒv‘ÎÛ‚Ì•¶š—ñ‚ğŠi”[iRubyj
            EscapeList = Array("$", "*", """")
        Case "block"
            'ƒGƒXƒP[ƒv‘ÎÛ‚Ì•¶š—ñ‚ğŠi”[iƒqƒAƒhƒLƒ…ƒƒ“ƒgj
            EscapeList = Array("""")
        Case Else
    End Select
    
    'ƒGƒXƒP[ƒv•¶š‚ğ•t‰Á
    For Each el In EscapeList
        strTargetVal = Replace(strTargetVal, el, "€" + el)
    Next
    
    SetEscape = strTargetVal
End Function


'------------------------------'
'ƒZƒ‹“à‰üsƒR[ƒh•ÏŠ·
'------------------------------'
Private Function SetLineBreak(ByVal strTarget As String) As String
      SetLineBreak = Replace(strTarget, vbLf, "€n")
End Function

               
'------------------------------'
'Yaml_Keyİ’è
'------------------------------'
Private Function SetYamlKey(ByVal strTargetSheet As String, ByVal strPatternType As String) As String
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)

    Select Case strPatternType
        Case "A", "B", "C", "D", "H", "J"
            strYamlKey = WS.Cells(6, 3).Value
            strYamlKey = Mid(strYamlKey, InStrRev(strYamlKey, "/") + 1)
            strYamlKey = Replace(strYamlKey, ".", "_")
            strYamlKey = Replace(strYamlKey, "-", "_")
            strYamlKey = StrConv(strYamlKey, vbNarrow + vbProperCase)

        Case "E"
            If strTargetSheet = "yƒOƒ‹[ƒvz" Then
                strYamlKey = "Group"
            ElseIf strTargetSheet = "yƒ†[ƒUz" Then
                strYamlKey = "User"
            ElseIf strTargetSheet = "yƒtƒ@ƒCƒ‹”z•zz" Then
                strYamlKey = "File"
            ElseIf strTargetSheet = "yƒfƒBƒŒƒNƒgƒŠì¬z" Then
                strYamlKey = "Directory"
            ElseIf strTargetSheet = "yƒpƒbƒP[ƒWz" Then
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
'TargetListƒtƒ@ƒCƒ‹ì¬
'------------------------------'
Private Function CreateTargetList()
    Dim i As Long
    
    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiOpenj
    Call OpenTextStream
    
    'ƒR[ƒh‘‚«o‚µ
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
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set ListWS = ThisWorkbook.Worksheets("ƒR[ƒhˆê——")
    
    'ŠJnˆÊ’uİ’è
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
            
            'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
            Set WS = ThisWorkbook.Worksheets("ƒŒƒVƒsˆê——")
            
            'ŠJnsİ’è
            i = 4
    
            'I—¹sİ’è
            EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
            
            'CookBook–¼æ“¾
            Do While i < EndRow
                If WS.Cells(i, 3).Value <> "" Then
                    strYamlCode = "     - '" + WS.Cells(i, 3).Value + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                End If
                i = i + 1
            Loop
            ActiveColumn = ActiveColumn + 1
        Loop
        '‰Šú‰»
        ActiveColumn = 4
        
        ActiveRow = ActiveRow + 1
    Loop
    
    '‘‚«o‚µŠÖ”ŒÄ‚Ño‚µ
    Call FileOutput("targetList.yml")
    
    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiClosej
    Call CloseTextStream

End Function

'------------------------------'
'ƒwƒbƒ_[ì¬
'------------------------------'
Private Function CreateHeader(ByVal strFileName As String)
    'ƒR[ƒh‘‚«o‚µ
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
'ƒpƒ^[ƒ“AƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternA(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yamlunamevkey–¼‚Ì—ñ‚ğæ“¾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'ƒRƒƒ“ƒgƒAƒEƒg•¶”»’è
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
                
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yamlunamevæ“¾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'ƒpƒ^[ƒ“BƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternB(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒuƒƒbƒN•¶AƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'ƒR[ƒh‘‚«o‚µilinesj
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µigrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "€n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'ƒR[ƒh‘‚«o‚µigrep_valj
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'ƒpƒ^[ƒ“CƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternC(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yamlunamevkey–¼‚Ì—ñ‚ğæ“¾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '•ÏX€–Úˆ—
    '------------------------------'
    '•ÏX€–Úˆ—ŠJnsİ’è
    ActiveRow = 6
    
    '•ÏX€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'ƒRƒƒ“ƒgƒAƒEƒg•¶”»’è
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yamlunamevæ“¾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '’Ç‰Á€–Úˆ—
    '------------------------------'
    '’Ç‰Á€–Úˆ—ŠJnsİ’è
    ActiveRow = ActiveRow + 3
    
    '’Ç‰Á€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi’Ç‰Á€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒuƒƒbƒN•¶AƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'ƒR[ƒh‘‚«o‚µilinesj
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µigrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "€n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'ƒR[ƒh‘‚«o‚µigrep_valj
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'ƒpƒ^[ƒ“DƒR[ƒh¶¬igrub.confj
'------------------------------'
Private Function CreatePatternD(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    '•Ï”éŒ¾
    Dim lines() As String
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yamlunamevkey–¼‚Ì—ñ‚ğæ“¾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '•ÏX€–Úˆ—
    '------------------------------'
    '•ÏX€–Úˆ—ŠJnsİ’è
    ActiveRow = 6
    
    '•ÏX€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'ƒRƒƒ“ƒgƒAƒEƒg•¶”»’è
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yamlunamevæ“¾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '’Ç‰Á€–Ú(kernels)
    '------------------------------'
    '’Ç‰Á€–Úˆ—ŠJnsİ’è
    ActiveRow = ActiveRow + 3
    
    '’Ç‰Á€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi’Ç‰Á€–Ú(kernels)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: 'kernel'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µigrep_valj
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)
    '------------------------------'
    '’Ç‰Á€–Úˆ—ŠJnsİ’è
    ActiveRow = ActiveRow + 3
    
    '’Ç‰Á€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi’Ç‰Á€–Ú(ƒtƒ@ƒCƒ‹––”ö)j", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒuƒƒbƒN•¶AƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            If strYamlPattern = "block" Then
                'ƒR[ƒh‘‚«o‚µilinesj
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "    :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µigrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "    :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "    :match_val: """ + strYamlValue + """" + "€n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'ƒR[ƒh‘‚«o‚µigrep_valj
                strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function

'------------------------------'
'ƒpƒ^[ƒ“EƒR[ƒh¶¬
'------------------------------'
Private Function CreatePatternE(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    Dim HashArray() As String
    Dim HashCount As Long
    Dim HashColumn As Long
    Dim HashRow As Long
    Dim strHashName As String
    Dim strHashValue As String
    Dim ha As Variant
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'HashŒÂ”æ“¾
    HashCount = 0
    HashColumn = 3
    Do While WS.Cells(5, HashColumn).Value <> ""
        HashCount = HashCount + 1
        HashColumn = HashColumn + 1
    Loop
    
    'Hash–¼ˆÊ’uæ“¾
    HashRow = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Row + 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine

    Select Case strTargetSheet
        Case "yƒ†[ƒUz"
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
                    'ƒR[ƒh‘‚«o‚µi-keyj
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlƒpƒ^[ƒ“”»’è
                    'ƒRƒƒ“ƒgƒAƒEƒg•¶‚©”»’è
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'ƒR[ƒh‘‚«o‚µipatternj
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '‰Šú‰»
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash”•ªsƒ‹[ƒvˆ—
                    Do While Count < HashCount
                        'ƒR[ƒh‘‚«o‚µiHash–¼j
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
        Case "yƒpƒbƒP[ƒWz"
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
                    'ƒR[ƒh‘‚«o‚µi-keyj
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlƒpƒ^[ƒ“”»’è
                    'ƒRƒƒ“ƒgƒAƒEƒg•¶‚©”»’è
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'ƒR[ƒh‘‚«o‚µipatternj
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '‰Šú‰»
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash”•ªsƒ‹[ƒvˆ—
                    Do While Count < HashCount
                        'ƒR[ƒh‘‚«o‚µiHash–¼j
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
            Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
                If WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
                    'ƒR[ƒh‘‚«o‚µi-keyj
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'yamlƒpƒ^[ƒ“”»’è
                    'ƒRƒƒ“ƒgƒAƒEƒg•¶‚©”»’è
                    If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                        strYamlPattern = "comment"
                    Else
                        strYamlPattern = "exist"
                    End If
                
                    'ƒR[ƒh‘‚«o‚µipatternj
                    strYamlCode = "    :pattern: '" + strYamlPattern + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    '‰Šú‰»
                    Count = 0
                    ActiveColumn = 3
                    HashColumn = WS.Cells.Find("Hash–¼", LookIn:=xlValues, LookAt:=xlWhole).Column
                    'Hash”•ªsƒ‹[ƒvˆ—
                    Do While Count < HashCount
                        'ƒR[ƒh‘‚«o‚µiHash–¼j
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
'ƒpƒ^[ƒ“FƒR[ƒh¶¬i©“®‹N“®ƒT[ƒrƒXj
'------------------------------'
Private Function CreatePatternF(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
        If WS.Cells(ActiveRow, 3) <> "" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            If WS.Cells(ActiveRow, ActiveFileColumn) = "" Or WS.Cells(ActiveRow, ActiveFileColumn) = "-" Then
                'ƒR[ƒh‘‚«o‚µipatternj
                strYamlCode = "    :pattern: 'noexist'"
                WriteStream.WriteText strYamlCode, adWriteLine
        
            ElseIf WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
                'ƒR[ƒh‘‚«o‚µipatternj
                strYamlCode = "    :pattern: 'exist'"
                WriteStream.WriteText strYamlCode, adWriteLine
            
                Count = 0
                Do While Count < 7
                    'ƒR[ƒh‘‚«o‚µirunlevelj
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
'ƒpƒ^[ƒ“GƒR[ƒh¶¬ixinetdƒT[ƒrƒXj
'------------------------------'
Private Function CreatePatternG(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 3) <> "ˆÈã"
        If WS.Cells(ActiveRow, 3) <> "" Then
    
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + WS.Cells(ActiveRow, 3) + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'ƒR[ƒh‘‚«o‚µipatternj
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
'ƒpƒ^[ƒ“HƒR[ƒh¶¬iHostsj
'------------------------------'
Private Function CreatePatternH(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 4) <> "ˆÈã"
        If WS.Cells(ActiveRow, 4) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            ActiveColumn = 5
            Do While ActiveColumn < 8
                If WS.Cells(ActiveRow, ActiveColumn) <> "" Then
                    'ƒR[ƒh‘‚«o‚µi-keyj
                    strYamlCode = "  - " + LCase(strYamlKey) + " :"
                    WriteStream.WriteText strYamlCode, adWriteLine
            
                    'ƒR[ƒh‘‚«o‚µiipaddressj
                    strYamlCode = "    :ipaddress: '" + WS.Cells(ActiveRow, 4) + "'"
                    WriteStream.WriteText strYamlCode, adWriteLine
                
                    'ƒR[ƒh‘‚«o‚µiHostname)
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
'ƒpƒ^[ƒ“IƒR[ƒh¶¬iƒ†[ƒUŠÂ‹«•Ï”j
'------------------------------'
Private Function CreatePatternI(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ŠJnsİ’è
    ActiveRow = 6
    
    'I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈã", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    'ƒR[ƒh‘‚«o‚µikeyj
    If strEvFlag <> strFileName Then
        strYamlCode = strYamlKey + ":"
        WriteStream.WriteText strYamlCode, adWriteLine
        strEvFlag = strFileName
    End If
    
    'ƒR[ƒh‘‚«o‚µi-keyj
    strYamlCode = "  - " + LCase(strYamlKey) + " :"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'ƒR[ƒh‘‚«o‚µipathj
    strYamlCode = "    :path: '" + WS.Cells(6, 5) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'ƒR[ƒh‘‚«o‚µifile_namej
    strYamlCode = "    :file_name: '" + WS.Cells(6, 6) + "'"
    WriteStream.WriteText strYamlCode, adWriteLine

    'ƒR[ƒh‘‚«o‚µivaluej
    strYamlCode = "    :value:"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    Do While WS.Cells(ActiveRow, 12) <> "ˆÈã"
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) = "›" Then
            'ƒR[ƒh‘‚«o‚µi-valj
            strYamlCode = "      - val :"
            WriteStream.WriteText strYamlCode, adWriteLine
    
            'ƒuƒƒbƒN•¶AƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If InStr(WS.Cells(ActiveRow, 12).Value, Chr(10)) <> 0 Then
                strYamlPattern = "block"
            ElseIf Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "        :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
                
            If strYamlPattern = "block" Then
                'ƒR[ƒh‘‚«o‚µilinesj
                Dim s() As String
                s = Split(WS.Cells(ActiveRow, 12).Value, vbLf)
                strYamlCode = "        :lines: '" & UBound(s) + 1 & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µigrep_val)
                strYamlValue = s(0)
                strYamlValue = SetEscape(strYamlValue, "grep")
                strYamlCode = "        :grep_val: '" & strYamlValue & "'"
                WriteStream.WriteText strYamlCode, adWriteLine
                
                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlValue = SetEscape(WS.Cells(ActiveRow, 12), "block")
                strYamlValue = SetLineBreak(strYamlValue)
                strYamlCode = "        :match_val: """ + strYamlValue + """" + "€n"
                WriteStream.WriteText strYamlCode, adWriteLine
            Else
                'ƒR[ƒh‘‚«o‚µigrep_valj
                strYamlCode = "        :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine

                'ƒR[ƒh‘‚«o‚µimatch_valj
                strYamlCode = "        :match_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "match") + "'"
                WriteStream.WriteText strYamlCode, adWriteLine
            End If

        End If
        
        ActiveRow = ActiveRow + 1
    Loop
End Function


'------------------------------'
'ƒpƒ^[ƒ“JƒR[ƒh¶¬isysctl.confj
'------------------------------'
Private Function CreatePatternJ(ByVal strTargetSheet As String, ByVal strFileName As String, ByVal strYamlKey As String)
    'ƒ[ƒNƒV[ƒgƒIƒuƒWƒFƒNƒgİ’è
    Set WS = ThisWorkbook.Worksheets(strTargetSheet)
     
    '‘ÎÛƒtƒ@ƒCƒ‹‚Ì—ñ‚ğİ’è
    ActiveFileColumn = WS.Cells.Find(strFileName, LookIn:=xlValues, LookAt:=xlWhole).Column
    
    'ƒR[ƒh‘‚«o‚µikeyj
    strYamlCode = strYamlKey + ":"
    WriteStream.WriteText strYamlCode, adWriteLine
    
    'Yamlunamevkey–¼‚Ì—ñ‚ğæ“¾
    ActiveColumn = WS.Cells.Find("Yaml_name", LookIn:=xlValues, LookAt:=xlWhole).Column
    
    '------------------------------'
    '•ÏX€–Úˆ—
    '------------------------------'
    '•ÏX€–Úˆ—ŠJnsİ’è
    ActiveRow = 6
    
    '•ÏX€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi•ÏX€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, ActiveFileColumn) <> "" And WS.Cells(ActiveRow, 12) <> "" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
        
            'ƒRƒƒ“ƒgƒAƒEƒg•¶”»’è
            If Left(LTrim(WS.Cells(ActiveRow, ActiveFileColumn).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'Yamlunamevæ“¾
            strYamlName = WS.Cells(ActiveRow, ActiveColumn).Value
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: '" + strYamlName + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
    '------------------------------'
    '’Ç‰Á€–Úˆ—
    '------------------------------'
    '’Ç‰Á€–Úˆ—ŠJnsİ’è
    ActiveRow = ActiveRow + 3
    
    '’Ç‰Á€–Úˆ—I—¹sİ’è
    EndRow = WS.Cells.Find("ˆÈãi’Ç‰Á€–Új", LookIn:=xlValues, LookAt:=xlWhole).Row - 1
    
    Do While ActiveRow < EndRow
        If WS.Cells(ActiveRow, 12) <> "" And WS.Cells(ActiveRow, ActiveFileColumn) <> "" Then
            'ƒR[ƒh‘‚«o‚µi-keyj
            strYamlCode = "  - " + LCase(strYamlKey) + " :"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒRƒƒ“ƒgƒAƒEƒg•¶‚Ì”»’è
            If Left(LTrim(WS.Cells(ActiveRow, 12).Value), 1) = "#" Then
                strYamlPattern = "comment"
            Else
                strYamlPattern = "exist"
            End If
            
            'ƒR[ƒh‘‚«o‚µipatternj
            strYamlCode = "    :pattern: '" + strYamlPattern + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µinamej
            strYamlCode = "    :name: 'add_parameter'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µigrep_valj
            strYamlCode = "    :grep_val: '" + SetEscape(WS.Cells(ActiveRow, 12), "grep") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
            'ƒR[ƒh‘‚«o‚µimatch_valj
            strYamlCode = "    :match_val: '" + SetEscape(WS.Cells(ActiveRow, ActiveFileColumn), "match") + "'"
            WriteStream.WriteText strYamlCode, adWriteLine
            
        End If
        ActiveRow = ActiveRow + 1
    Loop
    
End Function


Sub test()
    strTargetSheet = "y‹N“®ƒXƒNƒŠƒvƒgzrc.local"
    strFileName = "stprdb01"
    strYamlKey = "Rc_local"
    'ƒtƒHƒ‹ƒ_”»’è
    OutputPath = ThisWorkbook.Worksheets("ƒƒCƒ“").Range("C5").Value
    If CheckPath(OutputPath) = False Then
        MsgBox "o—ÍæƒtƒHƒ‹ƒ_‚ÉŒë‚è‚ª‚ ‚è‚Ü‚·B", vbCritical
        Exit Sub
    End If

    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiOpenj
    Call OpenTextStream
    Call CreatePatternC(strTargetSheet, strFileName, strYamlKey)
        
    '‘‚«o‚µŠÖ”ŒÄ‚Ño‚µ
    strFileName = strFileName + ".yml"
    Call FileOutput(strFileName)
        
    'ƒeƒLƒXƒgƒXƒgƒŠ[ƒ€‚ÌƒIƒuƒWƒFƒNƒg‚ğİ’è‚·‚éiClosej
    Call CloseTextStream

End Sub

