Attribute VB_Name = "CSVUtils_Test"
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
Sub test()
    Dim csvText(10) As String
    Dim csvExpected(10) As Variant
    Dim csvTextErr(10) As String
    Dim i As Long, r As Long, f As Long
    Dim csv As Collection
    Dim csva
    Dim csvs As String
    
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
      
    Debug.Print "----- Testing default error raise mode ----------------"
    
    ' In default, disable raising error
    ' one error for each function
    Err.Clear
    Set csv = ParseCSVToCollection(csvTextErr(0))
    If Not csv Is Nothing Or Err.Number <> 10002 Then Debug.Print "TEST NG 0a:" & Err.Number
    Err.Clear
    csva = ParseCSVToArray(csvTextErr(0))
    If Not IsNull(csva) Or Err.Number <> 10002 Then Debug.Print "TEST NG 0b:" & Err.Number
    Err.Clear
    Dim s As String
    csvs = ConvertArrayToCSV(s)
    If csvs <> "" Or Err.Number <> 10004 Then Debug.Print "TEST NG 0c:" & Err.Number
    Err.Clear
                
    Debug.Print "----- Testing error raise mode = AnyErrIsFatal ----------------"
    
    ' enabled raising error
    ' one error for each function
    Dim errorCnt As Long
    SetCSVUtilsAnyErrorIsFatal True 'enable
    On Error GoTo ErrCatch
    Set csv = ParseCSVToCollection(csvTextErr(0))
    csva = ParseCSVToArray(csvTextErr(0))
    csvs = ConvertArrayToCSV(s)
    GoTo NextTest
ErrCatch:
    errorCnt = errorCnt + 1
    If Err.Number <> 10002 And Err.Number <> 10004 Then Debug.Print "TEST NG 3:" & Err.Number
    Resume Next
NextTest:
    If errorCnt <> 3 Then Debug.Print "TEST NG 4:" & errorCnt
    On Error GoTo 0
        
    Debug.Print "----- Testing success data for parseXXXX() -------------------"
    For i = 0 To 9
        Set csv = ParseCSVToCollection(csvText(i))
        If csv Is Nothing Then Debug.Print "TEST NG"
        If Err.Number <> 0 Then Debug.Print "TEST NG"
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
        If Err.Number <> 0 Then Debug.Print "TEST NG"
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

    Debug.Print "----- Testing error data  for parseXXXX() ----------------"

    SetCSVUtilsAnyErrorIsFatal False 'disable
    
    Err.Clear
    Set csv = ParseCSVToCollection(csvTextErr(0))
    If Not csv Is Nothing Or Err.Number <> 10002 Then Debug.Print "TEST NG 0a:" & Err.Number
    Err.Clear
    csva = ParseCSVToArray(csvTextErr(0))
    If Not IsNull(csva) Or Err.Number <> 10002 Then Debug.Print "TEST NG 0b:" & Err.Number
    Err.Clear

    Set csv = ParseCSVToCollection(csvTextErr(1))
    If Not csv Is Nothing Or Err.Number <> 10003 Then Debug.Print "TEST NG 1a:" & Err.Number
    Err.Clear
    csva = ParseCSVToArray(csvTextErr(1))
    If Not IsNull(csva) Or Err.Number <> 10003 Then Debug.Print "TEST NG 1b:" & Err.Number
    Err.Clear
    
    Set csv = ParseCSVToCollection(csvTextErr(2))
    If Not csv Is Nothing Or Err.Number <> 10001 Then Debug.Print "TEST NG 2a:" & Err.Number
    Err.Clear
    csva = ParseCSVToArray(csvTextErr(2))
    If Not IsNull(csva) Or Err.Number <> 10001 Then Debug.Print "TEST NG 2b:" & Err.Number
    Err.Clear
    
    Debug.Print "----- Testing success data for ConvertArrayToCSV() -------------------"
    
    'including comma, double-quote, cr, lf, crlf, space
    s = "aaa , bbb,ccc" & vbCrLf & """x,xx"",""y""""yy"",""zz" & vbCr & "z""" & vbCrLf & """aa" & vbLf & "a"",""bb" & vbCrLf & "b"",ccc" & vbCrLf
    csvs = ConvertArrayToCSV(ParseCSVToArray(s))
    If Err.Number <> 0 Or csvs <> s Then Debug.Print "TEST NG 3a"
    'array range not starts with 1
    Dim aa1(0 To 1, 2 To 3) As String
    aa1(0, 2) = 1: aa1(1, 3) = 1
    csvs = ConvertArrayToCSV(aa1)
    If Err.Number <> 0 Or csvs <> "1," & vbCrLf & ",1" & vbCrLf Then Debug.Print "TEST NG 3b"
    Dim aa2(2 To 3, 0 To 1) As String
    aa2(2, 0) = 1: aa2(3, 1) = 1
    csvs = ConvertArrayToCSV(aa2)
    If Err.Number <> 0 Or csvs <> "1," & vbCrLf & ",1" & vbCrLf Then Debug.Print "TEST NG 3c"
    'Date type formatting
    Dim aa3(1 To 1, 1 To 2) As Variant
    aa3(1, 1) = DateSerial(2020, 1, 9)
    csvs = ConvertArrayToCSV(aa3)
    If Err.Number <> 0 Or csvs <> "2020/1/9," & vbCrLf Then Debug.Print "TEST NG 3d"
    csvs = ConvertArrayToCSV(aa3, "yyyy/mm/dd")
    If Err.Number <> 0 Or csvs <> "2020/01/09," & vbCrLf Then Debug.Print "TEST NG 3e"
    
    Debug.Print "----- Testing error data for ConvertArrayToCSV() -------------------"
    
    Err.Clear
    csvs = ConvertArrayToCSV(s)
    If csvs <> "" Or Err.Number <> 10004 Then Debug.Print "TEST NG 4a:" & Err.Number
    Err.Clear
    Dim a(2) As String
    csvs = ConvertArrayToCSV(a)
    If csvs <> "" Or Err.Number <> 9 Then Debug.Print "TEST NG 4a:" & Err.Number
    Err.Clear
    
    Debug.Print "----- End Testing ----------------"
    
End Sub


'
'  Performance TEST
'
Sub PerfTest()
  Dim flds(4) As String
  Dim csv As String, csv0 As String
  Dim i As Long, j As Long
  Dim t As Single
  Dim a As Variant
  
  csv = ""
  flds(0) = "abcdefg,"
  flds(1) = """hij,klmn"","
  flds(2) = """123""""456"","
  flds(3) = """opqrdtuv"","
  For j = 1 To 100 'columns
    csv = csv & flds(j Mod 4)
  Next
  csv = csv & vbCrLf
  For i = 1 To 13
    csv = csv & csv
  Next
  
  Debug.Print "START: " & Len(csv) & " Bytes"
  t = Timer
  'Call ParseCSVToCollection(csv)
  a = ParseCSVToArray(csv)
  If Err.Number <> 0 Then MsgBox Err.Number & Err.Source & Err.Description
  t = Timer - t
  Debug.Print "END: " & t & " sec."
  Debug.Print "records: " & UBound(a, 1) - 1
  Debug.Print "fields:  " & UBound(a, 2) - 1

End Sub
