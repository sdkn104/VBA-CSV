'VBScript Version
'  Error is always Fatal.
'  Array always starts with index 0

'Attribute VB_Name = "CSVUtils_Test"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
'
' License MIT (http://www.opensource.org/licenses/mit-license.php)
'
' This file is encoded by ShiftJIS (MS932?): ‚ ‚¢‚¤‚¦‚¨
'
'Option Explicit


Const IsVBA  = False

'
' Automatic TEST Procesure
'
'   If "End Testing" is shown in Immediate Window without "TEST NG", the TEST is pass.
'
Sub test()
    Dim csvText(10) 
    Dim csvExpected(10) 
    Dim csvTextErr(10) 
    Dim i , r , f 
    Dim csv 
    Dim csva
    Dim csvs 
    
    'error test data
    csvTextErr(0) = "aaa,""b""b"",ccc"  'illegal double quate
    csvTextErr(1) = "aaa,b""b,ccc"  'illegal field form (double quote in field)
    csvTextErr(2) = "aaa,bbb,ccc" & vbCrLf & "xxx,yyy" 'different field number
    
    ' success test data
    csvText(0) = ",aaa,SP111SP,‚ CRLF"""",""xxx"",SP""y,yy""SP,""SPz""""zCRLF""""zSP""SP"
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
    'For i = 0 To 3: MsgBox "[" & csvText(i) & "]": Next
    csvExpected(0) = Array(Array("", "aaa", "SP111SP", "‚ "), Array("", "xxx", "y,yy", "SPz""zCRLF""zSP"))
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
      
    If IsVBA Then
        MsgBox "----- Testing default error raise mode ----------------"
        
        ' In default, disable raising error
        ' one error for each function
        Err.Clear
        Set csv = ParseCSVToCollection(csvTextErr(0))
        If Not csv Is Nothing Or Err.Number <> 10002 Then MsgBox "TEST NG 0a:" & Err.Number
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(0))
        If Not IsNull(csva) Or Err.Number <> 10002 Then MsgBox "TEST NG 0b:" & Err.Number
        Err.Clear
        Dim s 
        csvs = ConvertArrayToCSV(s,"yyyy/m/d")
        If csvs <> "" Or Err.Number <> 10004 Then MsgBox "TEST NG 0c:" & Err.Number
        Err.Clear
    End If
                
    If IsVBA Then
        MsgBox "----- Testing error raise mode = AnyErrIsFatal ----------------"
        
        ' enabled raising error
        ' one error for each function
        Dim errorCnt 
        SetCSVUtilsAnyErrorIsFatal True 'enable
'        On Error GoTo ErrCatch
        Set csv = ParseCSVToCollection(csvTextErr(0))
        csva = ParseCSVToArray(csvTextErr(0))
        csvs = ConvertArrayToCSV(s,"yyyy/m/d")
'        GoTo NextTest
'ErrCatch:
        errorCnt = errorCnt + 1
        If Err.Number <> 10002 And Err.Number <> 10004 Then MsgBox "TEST NG 3:" & Err.Number
        'Resume Next
'NextTest:
        If errorCnt <> 3 Then MsgBox "TEST NG 4:" & errorCnt
'        On Error GoTo 0
    End If
    
    MsgBox "----- Testing success data for parseXXXX() -------------------"
    Dim arrStart 
    arrStart = 1
    If Not IsVBA Then arrStart = 0
    For i = 0 To 9
        Set csv = ParseCSVToCollection(csvText(i))
        If csv Is Nothing Then MsgBox "TEST NG"
        If Err.Number <> 0 Then MsgBox "TEST NG"
        If csv.Count <> UBound(csvExpected(i)) + 1 Then MsgBox "TEST NG row count"
        For r = 1 To csv.Count
          If csv.Item(r).Count <> UBound(csvExpected(i)(r - 1)) + 1 Then MsgBox "TEST NG col count"
          For f = 1 To csv.Item(r).Count
            If csv.Item(r).Item(f) <> csvExpected(i)(r - 1)(f - 1) Then MsgBox "TEST NG value"
            'MsgBox "[" & csv(r)(f) & "]"
          Next
        Next
        
        csva = ParseCSVToArray(csvText(i))
        If IsNull(csva) Then MsgBox "TEST NG"
        If Err.Number <> 0 Then MsgBox "TEST NG"
        If Not (LBound(csva, 1) = arrStart Or (LBound(csva, 1) = 0 And UBound(csva, 1) = -1)) Then MsgBox "TEST NG illegal array bounds"
        If UBound(csva, 1) - LBound(csva, 1) + 1 <> UBound(csvExpected(i)) + 1 Then MsgBox "TEST NG row count 2"
        For r = LBound(csva, 1) To UBound(csva, 1)
          If LBound(csva, 2) <> arrStart Or UBound(csva, 2) <> UBound(csvExpected(i)(r - arrStart)) + arrStart Then MsgBox "TEST NG col count 2"
          For f = LBound(csva, 2) To UBound(csva, 2)
            If csva(r, f) <> csvExpected(i)(r - arrStart)(f - arrStart) Then MsgBox "TEST NG value 2"
            'MsgBox "[" & csva(r, f) & "]"
            'MsgBox "[" & csvExpected(i)(r - 1)(f - 1) & "]"
          Next
        Next
    Next

    If IsVBA Then
        MsgBox "----- Testing error data  for parseXXXX() ----------------"
    
        SetCSVUtilsAnyErrorIsFatal False 'disable
        
        Err.Clear
        Set csv = ParseCSVToCollection(csvTextErr(0))
        If Not csv Is Nothing Or Err.Number <> 10002 Then MsgBox "TEST NG 0a:" & Err.Number
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(0))
        If Not IsNull(csva) Or Err.Number <> 10002 Then MsgBox "TEST NG 0b:" & Err.Number
        Err.Clear
    
        Set csv = ParseCSVToCollection(csvTextErr(1))
        If Not csv Is Nothing Or Err.Number <> 10002 Then MsgBox "TEST NG 1a:" & Err.Number
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(1))
        If Not IsNull(csva) Or Err.Number <> 10002 Then MsgBox "TEST NG 1b:" & Err.Number
        Err.Clear
        
        Set csv = ParseCSVToCollection(csvTextErr(2))
        If Not csv Is Nothing Or Err.Number <> 10001 Then MsgBox "TEST NG 2a:" & Err.Number
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(2))
        If Not IsNull(csva) Or Err.Number <> 10001 Then MsgBox "TEST NG 2b:" & Err.Number
        Err.Clear
    End If
    
    MsgBox "----- Testing success data for ConvertArrayToCSV() -------------------"
    
    'including comma, double-quote, cr, lf, crlf, space
    s = "aaa , bbb,ccc" & vbCrLf & """x,xx"",""y""""yy"",""zz" & vbCr & "z""" & vbCrLf & """aa" & vbLf & "a"",""bb" & vbCrLf & "b"",ccc" & vbCrLf
    csva = ParseCSVToArray(s)
    csvs = ConvertArrayToCSV(csva,"yyyy/m/d")
    If Err.Number <> 0 Or csvs <> s Then MsgBox "TEST NG 3a"
    If IsVBA Then
        'array range not starts with 1 'this is not needed for VBScript
        Dim aa1() 
        ReDim aa1(1-0, 3-2) 
        aa1(0, 2) = 1: aa1(1, 3) = 1
        csvs = ConvertArrayToCSV(aa1,"yyyy/m/d")
        If Err.Number <> 0 Or csvs <> "1," & vbCrLf & ",1" & vbCrLf Then MsgBox "TEST NG 3b"
        Dim aa2() 
        ReDim aa2(3-2, 1-0) 
        aa2(2, 0) = 1: aa2(3, 1) = 1
        csvs = ConvertArrayToCSV(aa2,"yyyy/m/d")
        If Err.Number <> 0 Or csvs <> "1," & vbCrLf & ",1" & vbCrLf Then MsgBox "TEST NG 3c"
    End If
    'Date type formatting
    Dim aa3(0, 1) 
    aa3(0, 0) = DateSerial(2020, 1, 9)
    csvs = ConvertArrayToCSV(aa3,"yyyy/m/d")
    If Err.Number <> 0 Or csvs <> "2020/1/9," & vbCrLf Then MsgBox "TEST NG 3d"
    csvs = ConvertArrayToCSV(aa3, "yyyy/mm/dd")
    If Err.Number <> 0 Or csvs <> "2020/01/09," & vbCrLf Then MsgBox "TEST NG 3e"
    
    If IsVBA Then
        MsgBox "----- Testing error data for ConvertArrayToCSV() -------------------"
        
        Err.Clear
        csvs = ConvertArrayToCSV(s,"yyyy/m/d")
        If csvs <> "" Or Err.Number <> 10004 Then MsgBox "TEST NG 4a:" & Err.Number
        Err.Clear
        Dim a(2) 
        csvs = ConvertArrayToCSV(a,"yyyy/m/d")
        If csvs <> "" Or Err.Number <> 9 Then MsgBox "TEST NG 4b:" & Err.Number
        Err.Clear
    End If
    
    MsgBox "----- End Testing ----------------"
    
End Sub


'
'  Performance TEST
'
Sub PerfTest()
  Dim flds(4) 
  Dim csv , csv0 
  Dim i , j 
  Dim t 
  Dim a 
  
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
  
  MsgBox "START parser: " & Len(csv) & " Bytes"
  t = Timer
  'Call ParseCSVToCollection(csv)
  a = ParseCSVToArray(csv)
  If Err.Number <> 0 Then MsgBox Err.Number & Err.Source & Err.Description
  t = Timer - t
  MsgBox "END: " & t & " sec."
  MsgBox " records: " & UBound(a, 1) - 1
  MsgBox " fields:  " & UBound(a, 2) - 1

  MsgBox "START writer:"
  t = Timer
  csv = ConvertArrayToCSV(a,"yyyy/m/d")
  If Err.Number <> 0 Then MsgBox Err.Number & Err.Source & Err.Description
  t = Timer - t
  MsgBox "END: " & t & " sec."


End Sub


