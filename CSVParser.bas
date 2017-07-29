Attribute VB_Name = "CSVParser"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
'
' License MIT (http://www.opensource.org/licenses/mit-license.php)
'
Option Explicit


Sub test()
    Dim csvText As String
    Dim csv As Collection
    Dim r As Collection
    Dim c As Variant
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.GetFile("C:\Users\sdkn1\Desktop\Book1.csv").OpenAsTextStream
        csvText = .ReadAll
        .Close
    End With
    Set FSO = Nothing
    
    'csvText = "a, b ,あ" & vbCrLf & "1,2, "" 3,""""3" & vbCrLf & "3 "" " & vbCrLf
        
    Set csv = ParseCSV(csvText)
    For Each r In csv
      Debug.Print "------"
      For Each c In r
        Debug.Print "[" & c & "]"
      Next
    Next
    Debug.Print "--------"
End Sub



Public Function ParseCSV(ByRef csvText As String) As Collection
    Dim csvLinesIdx As Long
    Dim csvLinesIdxMax As Long
    Dim lineText As String
    Dim recordText As String
    Dim fieldText As String
    Dim recLen As Long
    Dim regNL
    Dim regField
    Dim csvLines
    Dim cols
    Dim mField
    
    Set ParseCSV = New Collection  'as new Collection
        
    Set regNL = CreateObject("VBScript.RegExp")
    Set regField = CreateObject("VBScript.RegExp")
    
    regField.Pattern = "(\s*""(([^""]|"""")*)""\s*|([^,""]*)),"
    regField.Global = True
        
    'Split into lines (leaving line break codes)
    regNL.Pattern = "(\r\n|\r|\n)\s*$"
    csvText = regNL.Replace(csvText, "") 'delete line break code at EOF
    regNL.Pattern = "(\r\n|\r|\n)"
    regNL.Global = True
    csvText = regNL.Replace(csvText, "$1_^^^_")
    csvLines = Split(csvText, "_^^^_")
    csvLinesIdx = LBound(csvLines)
    csvLinesIdxMax = UBound(csvLines)
        
    'extract records and fields
    Do While GetOneRecord(csvLines, csvLinesIdx, csvLinesIdxMax, recordText)
        recLen = 0
        Set cols = New Collection
        For Each mField In regField.Execute(recordText & ",")
            recLen = recLen + mField.Length
            fieldText = regField.Replace(mField.Value, "$2")
            If fieldText = "" Then fieldText = regField.Replace(mField.Value, "$4")
            fieldText = Replace(fieldText, """""", """")
            cols.Add fieldText
        Next
        ParseCSV.Add cols
        If recLen <> Len(recordText) + 1 Then Err.Raise 998
    Loop
End Function


'
' Get the next one record into recordText from csvLines()
'
Private Function GetOneRecord(ByRef csvLines As Variant, ByRef csvLinesIdx As Long, ByRef csvLinesIdxMax As Long, ByRef recordText As String) As Boolean
    Dim regNL
    Set regNL = CreateObject("VBScript.RegExp")
    regNL.Pattern = "(\r\n|\r|\n)$"
        recordText = ""
        Do While csvLinesIdx <= csvLinesIdxMax
          recordText = recordText & csvLines(csvLinesIdx)
          csvLinesIdx = csvLinesIdx + 1
          If StrCount(recordText, """") Mod 2 = 0 Then
            recordText = regNL.Replace(recordText, "") 'remove the trailing line break code
            GetOneRecord = True
            Exit Function
          End If
        Loop
        If StrCount(recordText, """") Mod 2 = 1 Then Err.Raise 999
        GetOneRecord = False
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


