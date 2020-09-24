Attribute VB_Name = "PasteCsv"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Dim FILEPATH As String
    Dim FILENAME As String

    FILEPATH = "C:\Users\maca\Desktop\WORK\20171110\"
    FILENAME = "20171108_075940_HMAHM00__memory"

    Call PasteCsv(FILEPATH, FILENAME)

    
End Sub




Private Function PasteCsv(ByVal FILEPATH As String, ByVal FILENAME As String)
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + FILEPATH + FILENAME + ".csv", Destination:=Range("$A$1"))
        .Name = FILENAME
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.Name = FILENAME
End Function
