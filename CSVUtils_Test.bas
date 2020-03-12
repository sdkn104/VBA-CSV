Attribute VB_Name = "CSVUtils_Test"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
' License MIT (http://www.opensource.org/licenses/mit-license.php)
' This file is encoded by ShiftJIS (MS932?): あいうえお
'
Option Explicit


Const IsVBA As Boolean = True


' Execute All Tests
Sub TestAll()
    FunctionTest
    PerformanceTest
End Sub


'
' Automatic Functional TEST Procesure
'
'   If "End All Functional Testing" is shown in Immediate Window without "TEST FAIL" messages, the TEST is pass.
'
Sub FunctionTest()
    Dim csvText(10) As String
    Dim csvExpected(10) As Variant
    Dim csvTextErr(10) As String
    Dim i As Long, r As Long, f As Long
    Dim csv As Collection
    Dim csva
    Dim csvs As String, csvs2 As String
    Dim csvd As Object, csvd2 As Object
    
    'error test data
    csvTextErr(0) = "aaa,""b""b"",ccc"  'illegal double quate
    csvTextErr(1) = "aaa,b""b,ccc"  'illegal field form (double quote in field)
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
      
    Debug.Print "******** START Functional Testing (If error occurs, print TEST FAIL message.) ************"
    
    If IsVBA Then
        Debug.Print "----- Testing default error raise mode ----------------"
        
        ' In default, disable raising error
        ' one error for each function
        Err.Clear
        Set csv = ParseCSVToCollection(csvTextErr(0))
        MUST_BE_ERROR_OBJ csv, 10002, "0a:"
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(0))
        MUST_BE_ERROR_VAR csva, 10002, "0b:"
        Err.Clear
        Dim s As String
        csvs = ConvertArrayToCSV(s)
        MUST_BE_ERROR_STR csvs, 10004, "0c:"
        Err.Clear
        Set csvd = ParseCSVToDictionary(csvTextErr(0))
        MUST_BE_ERROR_OBJ csvd, 10002, "0d:"
        Err.Clear
    End If
                
                
    If IsVBA Then
        Debug.Print "----- Testing error raise mode = AnyErrIsFatal ----------------"
        
        ' enabled raising error
        ' one error for each function
        Dim errorCnt As Long
        SetCSVUtilsAnyErrorIsFatal True 'enable
        On Error GoTo ErrCatch
        Set csv = ParseCSVToCollection(csvTextErr(0))
        csva = ParseCSVToArray(csvTextErr(0))
        csvs = ConvertArrayToCSV(s)
        Set csvd = ParseCSVToDictionary(csvTextErr(0))
        GoTo NextTest
ErrCatch:
        errorCnt = errorCnt + 1
        If Err.Number <> 10002 And Err.Number <> 10004 Then Debug.Print "TEST FAILED 3:" & Err.Number
        Resume Next
NextTest:
        If errorCnt <> 4 Then Debug.Print "TEST FAILED 4:" & errorCnt
        On Error GoTo 0
    End If
    
    Debug.Print "----- Testing success data for parseXXXX() -------------------"
    Dim arrStart As Long
    arrStart = 1
    If Not IsVBA Then arrStart = 0
    For i = 0 To 9
        'ParseCSVToCollection
        Set csv = ParseCSVToCollection(csvText(i), False)
        MUST_BE_SUCCESS_OBJ csv, "success"
        MUST_BE csv.Count = UBound(csvExpected(i)) + 1, " wrong row count"
        For r = 1 To csv.Count
          MUST_BE csv.Item(r).Count = UBound(csvExpected(i)(r - 1)) + 1, "wrong col count"
          For f = 1 To csv.Item(r).Count
            MUST_BE csv.Item(r).Item(f) = csvExpected(i)(r - 1)(f - 1), "wrong value"
            'Debug.Print "[" & csv(r)(f) & "]"
          Next
        Next
        'ParseCSVToArray
        csva = ParseCSVToArray(csvText(i), False)
        MUST_BE_SUCCESS_VAR csva, "success2"
        MUST_BE (LBound(csva, 1) = arrStart Or (LBound(csva, 1) = 0 And UBound(csva, 1) = -1)), "illegal array bounds"
        MUST_BE Not UBound(csva, 1) - LBound(csva, 1) + 1 <> UBound(csvExpected(i)) + 1, "row count 2"
        For r = LBound(csva, 1) To UBound(csva, 1)
          MUST_BE LBound(csva, 2) = arrStart And UBound(csva, 2) = UBound(csvExpected(i)(r - arrStart)) + arrStart, "col count 2"
          For f = LBound(csva, 2) To UBound(csva, 2)
            MUST_BE csva(r, f) = csvExpected(i)(r - arrStart)(f - arrStart), "value 2"
            'Debug.Print "[" & csva(r, f) & "]"
            'Debug.Print "[" & csvExpected(i)(r - 1)(f - 1) & "]"
          Next
        Next
        ' ParseCSVToDictionary.  checking only success and colum count, since using ParseCSVToCollection for parsing
        Dim k
        Set csvd = ParseCSVToDictionary(csvText(i), 1, False)
        MUST_BE_SUCCESS_OBJ csvd, "success"
        For Each k In csvd.keys
          MUST_BE csvd.Item(k).Count = UBound(csvExpected(i)(0)) + 1, "wrong col count"
        Next
        ' GetFieldDictionary.  checking only success, since using ParseCSVToCollection for parsing
        Set csvd = GetFieldDictionary(csvText(i))
        MUST_BE_SUCCESS_OBJ csvd, "success"
    Next

    ' ParseCSVToDictionary()
    Set csvd = ParseCSVToDictionary("", 1, False) ' empty
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 0, " wrong row count"
    Set csvd = ParseCSVToDictionary("a,b,c", 1, False) ' header only
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 0, " wrong row count"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3", 1, False) 'header and one data
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 1, " wrong row count"
    MUST_BE csvd("1")(1) = 1 And csvd("1")(2) = 2 And csvd("1")(3) = 3, " wrong value"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3" & vbCrLf & "11,12,13", 1, False) ' two uniq data, keyCol=1
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 2, " wrong row count"
    MUST_BE csvd("1")(1) = 1 And csvd("1")(2) = 2 And csvd("1")(3) = 3, " wrong value"
    MUST_BE csvd("11")(1) = 11 And csvd("11")(2) = 12 And csvd("11")(3) = 13, " wrong value"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3" & vbCrLf & "11,12,13") ' two uniq data, keyCol=default
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 2, " wrong row count"
    MUST_BE csvd("1")(1) = 1 And csvd("1")(2) = 2 And csvd("1")(3) = 3, " wrong value"
    MUST_BE csvd("11")(1) = 11 And csvd("11")(2) = 12 And csvd("11")(3) = 13, " wrong value"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3" & vbCrLf & "1,12,13", 1, False) ' two duplicated data, keyCol=1
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 1, " wrong row count"
    MUST_BE csvd("1")(1) = 1 And csvd("1")(2) = 12 And csvd("1")(3) = 13, " wrong value"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3" & vbCrLf & "1,12,13", 2, False) ' two uniq data, keyCol=2
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 2, " wrong row count"
    MUST_BE csvd("2")(1) = 1 And csvd("2")(2) = 2 And csvd("2")(3) = 3, " wrong value"
    MUST_BE csvd("12")(1) = 1 And csvd("12")(2) = 12 And csvd("12")(3) = 13, " wrong value"
    Set csvd = ParseCSVToDictionary("a,b,c" & vbCrLf & "1,2,3" & vbCrLf & "1,12,13") ' 2 duplicated data, keyCol=default
    MUST_BE_SUCCESS_OBJ csvd, "success"
    MUST_BE csvd.Count = 1, " wrong row count"
        
    If IsVBA Then
        Debug.Print "----- Testing error data  for parseXXXX() ----------------"
    
        SetCSVUtilsAnyErrorIsFatal False 'disable
        
        Err.Clear
        Set csv = ParseCSVToCollection(csvTextErr(0))
        MUST_BE_ERROR_OBJ csv, 10002, "0a:"
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(0))
        MUST_BE_ERROR_VAR csva, 10002, "0b:"
        Err.Clear
        Set csvd = ParseCSVToDictionary(csvTextErr(0))
        MUST_BE_ERROR_OBJ csvd, 10002, "0c:"
        Err.Clear
    
        Set csv = ParseCSVToCollection(csvTextErr(1))
        MUST_BE_ERROR_OBJ csv, 10002, "1a:"
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(1))
        MUST_BE_ERROR_VAR csva, 10002, "1b:"
        Err.Clear
        Set csvd = ParseCSVToDictionary(csvTextErr(1))
        MUST_BE_ERROR_OBJ csvd, 10002, "1c:"
        Err.Clear
        
        Set csv = ParseCSVToCollection(csvTextErr(2))
        MUST_BE_ERROR_OBJ csv, 10001, "2a:"
        Err.Clear
        csva = ParseCSVToArray(csvTextErr(2))
        MUST_BE_ERROR_VAR csva, 10001, "2b:"
        Err.Clear
        Set csvd = ParseCSVToDictionary(csvTextErr(2))
        MUST_BE_ERROR_OBJ csvd, 10001, "2c:"
        Err.Clear
    
    End If
    
    Debug.Print "----- Testing success data for ConvertArrayToCSV() -------------------"
    
    'fields including comma, double-quote, cr, lf, crlf, space
    s = "aaa , bbb,ccc" & vbCrLf & """x,xx"",""y""""yy"",""zz" & vbCr & "z""" & vbCrLf & """aa" & vbLf & "a"",""bb" & vbCrLf & "b"",ccc" & vbCrLf
    csva = ParseCSVToArray(s, False)
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "3a"
    MUST_BE csvs = s, "3a2"
    If IsVBA Then
        'array range not starts with 1
        Dim aa1() As String
        ReDim aa1(0 To 1, 2 To 3) As String
        aa1(0, 2) = 1: aa1(1, 3) = 1
        csvs = ConvertArrayToCSV(aa1)
        MUST_BE_SUCCESS_STR csvs, "3b"
        MUST_BE csvs = "1," & vbCrLf & ",1" & vbCrLf, "3b"
        Dim aa2() As String
        ReDim aa2(2 To 3, 0 To 1) As String
        aa2(2, 0) = 1: aa2(3, 1) = 1
        csvs = ConvertArrayToCSV(aa2)
        MUST_BE_SUCCESS_STR csvs, "3c"
        MUST_BE csvs = "1," & vbCrLf & ",1" & vbCrLf, "3c"
    End If
    'Date type formatting
    Dim aa3(0, 1) As Variant
    aa3(0, 0) = DateSerial(2020, 1, 9)
    If IsVBA Then '---- omit argument
        csvs = ConvertArrayToCSV(aa3)
        MUST_BE_SUCCESS_STR csvs, "3d"
        MUST_BE csvs = "2020/1/9," & vbCrLf, "3d"
    End If
    csvs = ConvertArrayToCSV(aa3, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "3d"
    MUST_BE csvs = "2020/1/9," & vbCrLf, "3d"
    csvs = ConvertArrayToCSV(aa3, "yyyy/mm/dd", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "3e"
    MUST_BE csvs = "2020/01/09," & vbCrLf, "3e"
    'recordSeparator (line terminator)
    s = "aa,bb" & vbCrLf & "cc,dd" & vbCrLf
    csva = ParseCSVToArray(s, False)
    If IsVBA Then '---- omit arg
       csvs = ConvertArrayToCSV(csva)
        MUST_BE_SUCCESS_STR csvs, "3f"
        MUST_BE csvs = s, "3f"
    End If
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "3g"
    MUST_BE csvs = s, "3g"
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, "xxx")
    MUST_BE_SUCCESS_STR csvs, "3h"
    MUST_BE csvs = "aa,bbxxxcc,ddxxx", "3h"
    ' quoting
    s = "012,12.43,1e3," & vbCrLf & "aaa,""a,b"","""""""",""" & vbCr & """" & vbCrLf
    csva = ParseCSVToArray(s, False)
    If IsVBA Then '---- omit arg
        csvs = ConvertArrayToCSV(csva)
        MUST_BE_SUCCESS_STR csvs, "3i": MUST_BE csvs = s, "3i"
    End If
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "3j": MUST_BE csvs = s, "3j"
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", All, vbCrLf)
    s = """012"",""12.43"",""1e3"",""""" & vbCrLf & """aaa"",""a,b"","""""""",""" & vbCr & """" & vbCrLf
    MUST_BE_SUCCESS_STR csvs, "3k": MUST_BE csvs = s, "3k"
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", NONNUMERIC, vbCrLf)
    s = "012,12.43,1e3,""""" & vbCrLf & """aaa"",""a,b"","""""""",""" & vbCr & """" & vbCrLf
    MUST_BE_SUCCESS_STR csvs, "3l": MUST_BE csvs = s, "3l"
    
    If IsVBA Then
        Debug.Print "----- Testing error data for ConvertArrayToCSV() -------------------"
        
        Err.Clear
        csvs = ConvertArrayToCSV(s)
        MUST_BE_ERROR_STR csvs, 10004, "4a:"
        Err.Clear
        Dim a(2) As String
        csvs = ConvertArrayToCSV(a)
        MUST_BE_ERROR_STR csvs, 9, "4b:"
        Err.Clear
    End If
    
    Debug.Print "----- Testing allowVariableNumOfFields for parseXXXX() -------------------"
    s = "012,12.43,1e3," & vbCrLf & "aaa,ab,,ccc" & vbCrLf ' not variable data
    ' parseCSVToArray
    csva = ParseCSVToArray(s, False)
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    csva = ParseCSVToArray(s)
    csvs2 = ConvertArrayToCSV(csva)
    MUST_BE_SUCCESS_STR csvs, "5a": MUST_BE csvs = csvs2, "5a"
    csva = ParseCSVToArray(s, True)
    csvs2 = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "5b": MUST_BE csvs = csvs2, "5b"
    ' parseCSVToDictionary
    Set csvd = ParseCSVToDictionary(s, 1, False)
    Set csvd2 = ParseCSVToDictionary(s, 1)
    MUST_BE_SUCCESS_OBJ csvd, "5c"
    MUST_BE csvd.Count = csvd2.Count, "5c"
    For Each k In csvd.keys
        MUST_BE csvd(k).Count = csvd2(k).Count, "5cv"
        MUST_BE csvd(k)(1) = csvd2(k)(1), "5cv"
        MUST_BE csvd(k)(2) = csvd2(k)(2), "5cv"
        MUST_BE csvd(k)(3) = csvd2(k)(3), "5cv"
    Next
    Set csvd2 = ParseCSVToDictionary(s, 1, True)
    MUST_BE_SUCCESS_OBJ csvd, "5d"
    MUST_BE csvd.Count = csvd2.Count, "5d"
    For Each k In csvd.keys
        MUST_BE csvd(k).Count = csvd2(k).Count, "5dv"
        MUST_BE csvd(k)(1) = csvd2(k)(1), "5dv"
        MUST_BE csvd(k)(2) = csvd2(k)(2), "5dv"
        MUST_BE csvd(k)(3) = csvd2(k)(3), "5dv"
    Next
    
    s = "012,12.43,1e3" & vbCrLf & "aaa,ab,,ccc" & vbCrLf ' variable data
    ' ParseCSVToArray
    csva = ParseCSVToArray(s, True)
    MUST_BE_SUCCESS_VAR csva, "5e"
    csvs = ConvertArrayToCSV(csva, "yyyy/m/d", MINIMAL, vbCrLf)
    MUST_BE_SUCCESS_STR csvs, "5f": MUST_BE csvs = "012,12.43,1e3," & vbCrLf & "aaa,ab,,ccc" & vbCrLf, "5f"
    ' parseCSVToDictionary
    Set csvd = ParseCSVToDictionary(s, 1, True)
    MUST_BE_SUCCESS_OBJ csvd, "5g"
    MUST_BE csvd.Count = 1, "5h"
    MUST_BE csvd("aaa")(1) = "aaa" And csvd("aaa")(2) = "ab" And csvd("aaa")(3) = "" And csvd("aaa")(4) = "ccc", "5h"
        
    If IsVBA Then
        SetCSVUtilsAnyErrorIsFatal False 'disable
        Err.Clear
        ' ParseCSVToArray
        csva = ParseCSVToArray(s, False)
        MUST_BE_ERROR_VAR csva, 10001, "5i:"
        Err.Clear
        csva = ParseCSVToArray(s)
        MUST_BE_ERROR_VAR csva, 10001, "5j:"
        Err.Clear
        ' parseCSVToDictionary
        Set csvd = ParseCSVToDictionary(s, 1, False)
        MUST_BE_ERROR_OBJ csvd, 10001, "5k"
        Err.Clear
        Set csvd = ParseCSVToDictionary(s, 1)
        MUST_BE_ERROR_OBJ csvd, 10001, "5l"
        Err.Clear
    End If

    Debug.Print "----- Testing GetFieldDictionary() for success data -------------------"
    Set csvd = GetFieldDictionary("") ' empty
    MUST_BE_SUCCESS_OBJ csvd, "6a"
    MUST_BE csvd.Count = 0, "6b"
    Set csvd = GetFieldDictionary(vbCrLf) ' one blank field
    MUST_BE_SUCCESS_OBJ csvd, "6d"
    MUST_BE csvd.Count = 1, "6e"
    MUST_BE csvd("") = 1, "6f"
    Set csvd = GetFieldDictionary("a,b,c") ' three uniq fields
    MUST_BE_SUCCESS_OBJ csvd, "6g"
    MUST_BE csvd.Count = 3, "6h"
    MUST_BE csvd("a") = 1 And csvd("b") = 2 And csvd("c") = 3, "6i"
    Set csvd = GetFieldDictionary("a,a,c") ' three duplicated fields
    MUST_BE_SUCCESS_OBJ csvd, "6j"
    MUST_BE csvd.Count = 2, "6k"
    MUST_BE csvd("a") = 2 And csvd("c") = 3, "6l"
    
    Debug.Print "----- Testing error data  for GetFieldDictionary() ----------------"

    SetCSVUtilsAnyErrorIsFatal False 'disable
    Err.Clear
    Set csvd = GetFieldDictionary(csvTextErr(0))
    MUST_BE_ERROR_OBJ csvd, 10002, "0d:"
    Err.Clear
    Set csvd = GetFieldDictionary(csvTextErr(1))
    MUST_BE_ERROR_OBJ csvd, 10002, "1d:"
    Err.Clear
    
    Debug.Print "******** End All Functional Testing ********"
    
End Sub


'
'  Performance TEST
'
Sub PerformanceTest()
  Dim flds(4) As String
  Dim csv As String, csv0 As String
  Dim i As Long, j As Long
  Dim t As Single
  Dim a As Variant
  
  Debug.Print "******** Start Perforance Test ********"
  
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
  
  Debug.Print "START parser: " & Len(csv) & " Bytes ..."
  t = Timer
  'Call ParseCSVToCollection(csv)
  a = ParseCSVToArray(csv, False)
  If Err.Number <> 0 Then MsgBox Err.Number & Err.Source & Err.Description
  t = Timer - t
  Debug.Print "END: " & t & " sec."
  Debug.Print " Data Size: " & UBound(a, 2) - 1 & " fields x " & UBound(a, 1) - 1 & " records"

  Debug.Print "START writer ..."
  t = Timer
  csv = ConvertArrayToCSV(a, "yyyy/m/d", MINIMAL, vbCrLf)
  If Err.Number <> 0 Then MsgBox Err.Number & Err.Source & Err.Description
  t = Timer - t
  Debug.Print "END: " & t & " sec."

  Debug.Print "******** End Performance Test ********"

End Sub



Sub MUST_BE_ERROR_OBJ(returned, errNumber As Long, msgText)
    MUST_BE returned Is Nothing And Err.Number = errNumber, msgText
End Sub

Sub MUST_BE_ERROR_VAR(returned, errNumber As Long, msgText)
    MUST_BE IsNull(returned) And Err.Number = errNumber, msgText
End Sub

Sub MUST_BE_ERROR_STR(returned, errNumber As Long, msgText)
    MUST_BE returned = "" And Err.Number = errNumber, msgText
End Sub

Sub MUST_BE_SUCCESS_OBJ(returned, msgText)
    MUST_BE Not returned Is Nothing And Err.Number = 0, msgText
End Sub

Sub MUST_BE_SUCCESS_VAR(returned, msgText)
    MUST_BE Not IsNull(returned) And Err.Number = 0, msgText
End Sub

Sub MUST_BE_SUCCESS_STR(returned, msgText)
    MUST_BE returned <> "" And Err.Number = 0, msgText
End Sub

Sub MUST_BE(cond, msgText)
    If Not cond Then Debug.Print "TEST FAILED " & msgText & Err.Number
End Sub


