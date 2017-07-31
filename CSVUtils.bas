Attribute VB_Name = "CSVUtils"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
'
' License MIT (http://www.opensource.org/licenses/mit-license.php)
'
Option Explicit


'----- ERROR HANDLER -------------
Public ParseCSVEnableRaiseError As Boolean  'default False




'----- ERROR HANDLER -----------------------------------

Private Sub ErrorHandler(code As Long, src As String, msg As String)
'  If ParseCSVEnableRaiseError Then
'    Err.raise code, src, msg
'  End If
'  On Error Resume Next
  If err.Number = 0 Then err.Raise code, src, msg
End Sub

Public Sub SetParseCSVEnableRaiseError(ByRef value As Boolean)
  ParseCSVEnableRaiseError = value
End Sub


'------ Public Function/Sub ---------------------------

Public Function ParseCSVToCollection(ByRef csvText As String) As Collection
    err.Clear
    If ParseCSVEnableRaiseError Then GoTo Head
    On Error Resume Next
Head:
    Dim csvLinesIdx As Long
    Dim csvLinesIdxMax As Long
    Dim csvTextTmp As String
    Dim lineText As String
    Dim recordText As String
    Dim fieldText As String
    Dim recLen As Long
    Dim regNL 'RegExp
    Dim regField 'RegExp
    Dim mField 'Match
    Dim csvLines As Variant
    Dim fields As Collection
    Dim csvCollection As Collection
    Set csvCollection = New Collection 'empty collection
    
    Set ParseCSVToCollection = csvCollection
        
    Set regNL = CreateObject("VBScript.RegExp")
    Set regField = CreateObject("VBScript.RegExp")
    
    regField.Pattern = "(\s*""(([^""]|"""")*)""\s*|([^,""]*)),"
    regField.Global = True
        
    'for empty text
    If csvText = "" Then Exit Function 'return empty collection
    
    'Split into lines (leaving line break codes)
    regNL.Pattern = "(\r\n|\r|\n)$"
    csvTextTmp = regNL.Replace(csvText, "") 'delete line break code at EOF
    regNL.Pattern = "(\r\n|\r|\n)"
    regNL.Global = True
    csvTextTmp = regNL.Replace(csvTextTmp, "$1_^`~_")
    csvLines = Split(csvTextTmp, "_^`~_")
    If csvTextTmp = "" Then csvLines = Array("") 'since VBA Split() returns empty(zero length) array for ""
    csvLinesIdx = LBound(csvLines)
    csvLinesIdxMax = UBound(csvLines)
    csvTextTmp = "" 'to free memory

    'extract records and fields
    Do While GetOneRecord(csvLines, csvLinesIdx, csvLinesIdxMax, recordText)
        recLen = 0
        Set fields = New Collection
        For Each mField In regField.Execute(recordText & ",")
            recLen = recLen + Len(mField.value)
            fieldText = regField.Replace(mField.value, "$2")
            If fieldText = "" Then fieldText = regField.Replace(mField.value, "$4")
            fieldText = Replace(fieldText, """""", """") 'un-escape
            fields.Add fieldText
        Next
        csvCollection.Add fields
        
        If csvCollection(1).Count <> fields.Count Then
            ErrorHandler 10001, "ParseCSVToCollection", "Syntax Error in CSV: numbers of fields are different among records"
            GoTo ErrorExit
        End If
        If recLen <> Len(recordText) + 1 Then
            ErrorHandler 10003, "ParseCSVToCollection", "Syntax Error in CSV: illegal field form"
            GoTo ErrorExit
        End If
    Loop
    If err.Number <> 0 Then GoTo ErrorExit
    
    Set ParseCSVToCollection = csvCollection
    Exit Function

ErrorExit:
    Set ParseCSVToCollection = Nothing
End Function


Public Function ParseCSVToArray(ByRef csvText As String) As Variant
    err.Clear
    If ParseCSVEnableRaiseError Then GoTo Head
    On Error Resume Next
Head:
    Dim csv As Collection
    Dim recCnt As Long, fldCnt As Long
    Dim csvArray() As String
    Dim ri As Long, fi As Long
    
    ParseCSVToArray = Null
  
    Set csv = ParseCSVToCollection(csvText)
    If csv Is Nothing Then  'error occur
        Exit Function
    End If
    
    recCnt = csv.Count
    If recCnt = 0 Then
        ParseCSVToArray = Split("", "/") 'return empty(zero length) String array of bound 0 TO -1
                                         '(https://msdn.microsoft.com/ja-jp/library/office/gg278528.aspx)
        Exit Function
    End If
    fldCnt = csv(1).Count
    
    ReDim csvArray(1 To recCnt, 1 To fldCnt) As String
    For ri = 1 To recCnt
      For fi = 1 To fldCnt
        csvArray(ri, fi) = csv(ri)(fi)
      Next
    Next
    
    ParseCSVToArray = csvArray
End Function



' ------------- Private function/sub ----------------------------------------

'
' Get the next one record from csvLines, and put it into recordText
'
Private Function GetOneRecord(ByRef csvLines As Variant, ByRef csvLinesIdx As Long, ByRef csvLinesIdxMax As Long, ByRef recordText As String) As Boolean
    Dim dQuateCnt As Long
    Dim lineText As String
    Dim regNL
    Set regNL = CreateObject("VBScript.RegExp")
    regNL.Pattern = "(\r\n|\r|\n)$"
    
    recordText = ""
    dQuateCnt = 0
    Do While csvLinesIdx <= csvLinesIdxMax
        lineText = csvLines(csvLinesIdx)
        recordText = recordText & lineText
        dQuateCnt = dQuateCnt + StrCount(lineText, """")
        csvLinesIdx = csvLinesIdx + 1
        If dQuateCnt Mod 2 = 0 Then  'if the number of double-quates is even, then the current field ends
            recordText = regNL.Replace(recordText, "") 'remove the trailing line break code
            GetOneRecord = True
            Exit Function
        End If
    Loop
    
    GetOneRecord = False
    If recordText <> "" Then
      ErrorHandler 10002, "ParseCSVToCollection", "Syntax Error in CSV: illegal double-quote code"
    End If
End Function


' count the string Target in Source
Private Function StrCount(Source As String, Target As String) As Long
    Dim n As Long, cnt As Long
    n = 0
    cnt = 0
    Do
        n = InStr(n + 1, Source, Target)
        If n = 0 Then
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop
    StrCount = cnt
End Function

