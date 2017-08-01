Attribute VB_Name = "Test"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
'
' License MIT (http://www.opensource.org/licenses/mit-license.php)
'
Option Explicit

'
' Automatic TEST Procesure
'
'   If "End Testing" is shown in Immediate Window without "TEST NG", the TEST is pass.
'
Sub Test()
    Dim csvText(10) As String
    Dim csvExpected(10) As Variant
    Dim csvTextErr(10) As String
    Dim i As Long, r As Long, f As Long
    Dim csv As Collection
    Dim csva
    
    'error test data
    csvTextErr(0) = "aaa,""b""b"",ccc"  'illegal double quate
    csvTextErr(1) = "aaa,b""""b,ccc"  'illegal field form (double quate in field)
    csvTextErr(2) = "aaa,bbb,ccc" & vbCrLf & "xxx,yyy" 'different field number
    
    ' success test data
    csvText(0) = ",aaa,SP111SP,あCRLF"""",""xxx"",SP""y,yy""SP,""SPz""""zCRLF""""zSP""SP"
    csvText(0) = Replace(csvText(0), "SP", " " & vbTab)
    csvText(0) = Replace(csvText(0), "CRLF", vbCrLf) ' no line break at EOF
    csvText(1) = csvText(0) & vbCrLf ' line break at EOF
    csvText(2) = Replace(csvText(1), vbCrLf, vbCr) ' CR
    csvText(3) = Replace(csvText(1), vbCrLf, vbLf) ' LF
    csvText(4) = "" 'empty
    csvText(5) = vbTab 'one record containing one TAB field
    csvText(6) = "," ' one record containing two blank field
    csvText(7) = vbCrLf ' one record containing one blank field
    csvText(8) = vbCrLf & vbCrLf ' two records containing one blank field
    csvText(9) = vbCrLf & vbTab ' two records containing one blank field, one TAB field
    'For i = 0 To 3: Debug.Print "[" & csvText(i) & "]": Next
    csvExpected(0) = Array(Array("", "aaa", "SP111SP", "あ"), Array("", "xxx", "y,yy", "SPz""zCRLF""zSP"))
    csvExpected(1) = Array(Array("", "", "", ""), Array("", "", "", ""))
    csvExpected(2) = Array(Array("", "", "", ""), Array("", "", "", ""))
    csvExpected(3) = Array(Array("", "", "", ""), Array("", "", "", ""))
    For r = LBound(csvExpected(0)) To UBound(csvExpected(0))
      For f = LBound(csvExpected(0)(r)) To UBound(csvExpected(0)(r))
        csvExpected(0)(r)(f) = Replace(csvExpected(0)(r)(f), "SP", " " & vbTab)
        csvExpected(0)(r)(f) = Replace(csvExpected(0)(r)(f), "CRLF", vbCrLf)
        csvExpected(1)(r)(f) = csvExpected(0)(r)(f)
        csvExpected(2)(r)(f) = Replace(csvExpected(1)(r)(f), vbCrLf, vbCr)
        csvExpected(3)(r)(f) = Replace(csvExpected(1)(r)(f), vbCrLf, vbLf)
      Next
    Next
    csvExpected(4) = Array()
    csvExpected(5) = Array(Array(vbTab))
    csvExpected(6) = Array(Array("", ""))
    csvExpected(7) = Array(Array(""))
    csvExpected(8) = Array(Array(""), Array(""))
    csvExpected(9) = Array(Array(""), Array(vbTab))
      
    Debug.Print "----- Testing error raise mode ----------------"
    
    ' In default, disable raising error
    err.Clear
    Set csv = ParseCSVToCollection(csvTextErr(0))
    If Not csv Is Nothing Or err.Number <> 10002 Then Debug.Print "TEST NG 0a:" & err.Number
    err.Clear
    csva = ParseCSVToArray(csvTextErr(0))
    If Not IsNull(csva) Or err.Number <> 10002 Then Debug.Print "TEST NG 0b:" & err.Number
    err.Clear
            
    ' enabled raising error
    Dim errorCnt As Long
    SetParseCSVEnableRaiseError True 'enable
    On Error GoTo ErrCatch
    Set csv = ParseCSVToCollection(csvTextErr(0))
    csva = ParseCSVToArray(csvTextErr(0))
    GoTo NextTest
ErrCatch:
    errorCnt = errorCnt + 1
    If err.Number <> 10002 Then Debug.Print "TEST NG 3:" & err.Number
    Resume Next
NextTest:
    If errorCnt <> 2 Then Debug.Print "TEST NG 4:" & errorCnt
    On Error GoTo 0
        
    Debug.Print "----- Testing success data -------------------"
    For i = 0 To 9
        Set csv = ParseCSVToCollection(csvText(i))
        If csv Is Nothing Then Debug.Print "TEST NG"
        If err.Number <> 0 Then Debug.Print "TEST NG"
        If csv.Count <> UBound(csvExpected(i)) + 1 Then Debug.Print "TEST NG row count"
        For r = 1 To csv.Count
          If csv(r).Count <> UBound(csvExpected(i)(r - 1)) + 1 Then Debug.Print "TEST NG col count"
          For f = 1 To csv(r).Count
            If csv(r)(f) <> csvExpected(i)(r - 1)(f - 1) Then Debug.Print "TEST NG value"
            'Debug.Print "[" & csv(r)(f) & "]"
          Next
        Next
        
        csva = ParseCSVToArray(csvText(i))
        If IsNull(csva) Then Debug.Print "TEST NG"
        If err.Number <> 0 Then Debug.Print "TEST NG"
        If Not (LBound(csva, 1) = 1 Or (LBound(csva, 1) = 0 And UBound(csva, 1) = -1)) Then Debug.Print "TEST NG illegal array bounds"
        If UBound(csva, 1) - LBound(csva, 1) + 1 <> UBound(csvExpected(i)) + 1 Then Debug.Print "TEST NG row count"
        For r = LBound(csva, 1) To UBound(csva, 1)
          If LBound(csva, 2) <> 1 Or UBound(csva, 2) <> UBound(csvExpected(i)(r - 1)) + 1 Then Debug.Print "TEST NG col count"
          For f = LBound(csva, 2) To UBound(csva, 2)
            If csva(r, f) <> csvExpected(i)(r - 1)(f - 1) Then Debug.Print "TEST NG value"
            'Debug.Print "[" & csva(r, f) & "]"
            'Debug.Print "[" & csvExpected(i)(r - 1)(f - 1) & "]"
          Next
        Next
    Next

    Debug.Print "----- Testing error data ----------------"

    SetParseCSVEnableRaiseError False 'disable
    
    err.Clear
    Set csv = ParseCSVToCollection(csvTextErr(0))
    If Not csv Is Nothing Or err.Number <> 10002 Then Debug.Print "TEST NG 0a:" & err.Number
    err.Clear
    csva = ParseCSVToArray(csvTextErr(0))
    If Not IsNull(csva) Or err.Number <> 10002 Then Debug.Print "TEST NG 0b:" & err.Number
    err.Clear

    Set csv = ParseCSVToCollection(csvTextErr(1))
    If Not csv Is Nothing Or err.Number <> 10003 Then Debug.Print "TEST NG 1a:" & err.Number
    err.Clear
    csva = ParseCSVToArray(csvTextErr(1))
    If Not IsNull(csva) Or err.Number <> 10003 Then Debug.Print "TEST NG 1b:" & err.Number
    err.Clear
    
    Set csv = ParseCSVToCollection(csvTextErr(2))
    If Not csv Is Nothing Or err.Number <> 10001 Then Debug.Print "TEST NG 2a:" & err.Number
    err.Clear
    csva = ParseCSVToArray(csvTextErr(2))
    If Not IsNull(csva) Or err.Number <> 10001 Then Debug.Print "TEST NG 2b:" & err.Number
    err.Clear
    
    Debug.Print "----- End Testing ----------------"
            
End Sub

