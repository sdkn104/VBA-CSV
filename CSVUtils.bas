Attribute VB_Name = "CSVUtils"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
' License MIT (http://www.opensource.org/licenses/mit-license.php)
' Document: https://github.com/sdkn104/VBA-CSV/README.md
'
Option Explicit


' Variables used in FindNextSeparator()
Private nextSep1 As Long
Private nextSep2 As Long
Private nextSep3 As Long

'----- Enum -------------------------------------------------------------------------

' Field Quoting
'   Used for the argument 'quoting' of ConvertArrayToCSV()
'   This argument controls what kind of fields to be quoted
Public Enum CSVUtilsQuote
    MINIMAL = 0     ' quote the fields that requires quotation (i.e., that includes comma, return code, quotation mark)
    All = 1         ' quote all the fields
    NONNUMERIC = 2  ' quote non-numeric (Not IsNumeric()) fields
End Enum

'----- Global variables -------------------------------------------------------------

Private CSVUtilsAnyErrorIsFatal As Boolean  'default False

'----- ERROR HANDLER ----------------------------------------------------------------

'
' Error function
'
Private Sub ErrorRaise(code As Long, src As String, msg As String)
  ' raise only if this is the first error
  If Err.Number = 0 Then Err.Raise code, src, msg
End Sub

'
' Setting error handling mode
'
'  False (default) --- When run-time error occurs, the parser function returns special value (Nothing,  Null, etc.),
'                      and the error information is set to properties of Err object.
'  True            --- Any run-time error that occurs is fatal (an error message is displayed and execution stops).
'
Public Sub SetCSVUtilsAnyErrorIsFatal(ByRef value As Boolean)
  CSVUtilsAnyErrorIsFatal = value
End Sub


'------ Public Function/Sub --------------------------------------------------------

'
' Parse CSV text and retern Collection
'
'   Return a Collection of records; record is a Collection of fields
'   When error, return Nothing
'
Public Function ParseCSVToCollection(ByRef csvText As String, Optional ByRef allowVariableNumOfFields As Boolean = False, Optional ByRef headerOnly As Boolean = False) As Collection
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csvPos As Long
    Dim fieldText As String
    Dim nextSep As Long, nextSepType As Long, quoteCount As Long, fieldStart As Long, fieldLen As Long
    Dim fields As Collection
    Dim csvCollection As Collection
    Set csvCollection = New Collection 'empty collection
    
    Set ParseCSVToCollection = csvCollection
    
    'for empty text
    If csvText = "" Then Exit Function 'return empty collection
    
    ' Add trailing record separator if not
    If Right(csvText, 1) <> "" & vbCr And Right(csvText, 1) <> "" & vbLf Then
        csvText = csvText & vbCrLf
    End If
    
    'extract records and fields
    csvPos = 1
    Set fields = New Collection
    Call FindNextSeparatorInit(csvText)
    Do While FindNextSeparator(csvText, csvPos, fieldStart, fieldLen, nextSepType, quoteCount)
        fieldText = Mid(csvText, fieldStart, fieldLen)
        If Err.Number <> 0 Then Exit Do
        
        If quoteCount > 0 Then ' the field includes " (double-quote)
            fieldText = TrimQuotes(fieldText) 'get internal of ""
            If quoteCount > 2 Then 'the field includes double-quote in internal of ""
                fieldText = Replace(fieldText, """""", """") 'un-escape double quote
                If fieldText Like "=""*""" Then fieldText = Mid(fieldText, 3, Len(fieldText) - 3) 'remove MS quote (="...")
            End If
        End If
        'add to collection
        fields.Add fieldText
            
        If nextSepType <> 1 Then ' end of the record
            csvCollection.Add fields
            If headerOnly Then Exit Do
            If Not allowVariableNumOfFields And csvCollection.Item(1).Count <> fields.Count Then
                ErrorRaise 10001, "ParseCSVToCollection", "Syntax Error in CSV: numbers of fields are different among records"
                GoTo ErrorExit
            End If
            Set fields = New Collection
        End If
    Loop
    If Err.Number <> 0 Then GoTo ErrorExit
    
    Set ParseCSVToCollection = csvCollection
    Exit Function

ErrorExit:
    Set ParseCSVToCollection = Nothing
End Function


'
' Parse CSV text and return 2-dim array
'
'  Return 2-dim array --- String(1 TO recordCount, 1 TO fieldCount)
'  When CSV text is "", return empty array --- String(0 TO -1)
'  When error, return Null
'
Public Function ParseCSVToArray(ByRef csvText As String, Optional ByRef allowVariableNumOfFields As Boolean = False) As Variant
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csv As Collection
    Dim rowCount As Long, colCount As Long
    Dim csvArray() As String
    Dim ri As Long, fi As Long
    Dim sepIndex As Long
    Dim fieldStart As Long, fieldLen As Long, nextSepType As Long, quoteCount As Long
    Dim fieldText As String
    
    ParseCSVToArray = Null 'for error
  
    Dim sepArray1() As Long
    Dim sepArray2() As Long
    Dim sepArray3() As Long
    Dim sepArray4() As Long
    ReDim sepArray1(Len(csvText) / 40 + 64)
    ReDim sepArray2(Len(csvText) / 40 + 64)
    ReDim sepArray3(Len(csvText) / 40 + 64)
    ReDim sepArray4(Len(csvText) / 40 + 64)
  
    ' Parse CSV and get row/col count, sepArray1234
    Call ParseCSV(rowCount, colCount, sepArray1, sepArray2, sepArray3, sepArray4, csvText, allowVariableNumOfFields)
    If Err.Number <> 0 Then     'error occur
        Exit Function
    End If
    
    ' empty
    If rowCount = 0 Then
        ParseCSVToArray = Split("", "/") 'return empty(zero length) String array of bound 0 TO -1
                                         '(https://msdn.microsoft.com/ja-jp/library/office/gg278528.aspx)
        Exit Function
    End If
        
    ' allocate result array
    ReDim csvArray(1 To rowCount, 1 To colCount) As String
    
    ' fill result array
    sepIndex = 0
    ri = 1
    fi = 1
    Do
        fieldStart = sepArray1(sepIndex)
        If fieldStart = 0 Then Exit Do ' EOF
        
        fieldLen = sepArray2(sepIndex)
        nextSepType = sepArray3(sepIndex)
        quoteCount = sepArray4(sepIndex)
        fieldText = Mid(csvText, fieldStart, fieldLen)
        If quoteCount > 0 Then ' the field includes " (double-quote)
            fieldText = TrimQuotes(fieldText) 'get internal of ""
            If quoteCount > 2 Then 'the field includes double-quote in internal of ""
                fieldText = Replace(fieldText, """""", """") 'un-escape double quote
                If fieldText Like "=""*""" Then fieldText = Mid(fieldText, 3, Len(fieldText) - 3) 'remove MS quote (="...")
            End If
        End If
        csvArray(ri, fi) = fieldText
        fi = fi + 1
        If nextSepType <> 1 Then ' end of record
            ri = ri + 1
            fi = 1
        End If
        sepIndex = sepIndex + 1
    Loop
    ParseCSVToArray = csvArray
End Function


'
' Convert 2-dim array to CSV text string
'
'  inArray : 2-dim array of arbitary size/range and type.
'  fmtDate : format used for conversion from type Date to type String
'  When error, return ""
'
Public Function ConvertArrayToCSV(inArray As Variant, Optional fmtDate As String = "yyyy/m/d", _
                          Optional ByVal quoting As CSVUtilsQuote = CSVUtilsQuote.MINIMAL, _
                          Optional ByVal recordSeparator As String = vbCrLf) As String
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csv As String
    Dim r As Long, c As Long, ub2 As Long
    Dim v As Variant
    Dim cell As String
    Dim arrRecord As Variant, arrField As Variant
    
    'error check
    If Not IsArray(inArray) Then
        ErrorRaise 10004, "ConvertArrayToCSV", "Input argument inArray is not array"
        GoTo ErrorExit
    End If
    ub2 = UBound(inArray, 2)
    If Err.Number <> 0 Then 'expecting Err.Number = 9, Err.Description = "Subscript out of range", for inArray is 1-dim
        GoTo ErrorExit
    End If

    Dim rc As Long, cc As Long
    ReDim arrRecord(LBound(inArray, 1) To UBound(inArray, 1)) As String 'temporary array
    ReDim arrField(LBound(inArray, 2) To UBound(inArray, 2)) As String 'temporary array
    
    For r = LBound(inArray, 1) To UBound(inArray, 1)
      For c = LBound(inArray, 2) To UBound(inArray, 2)
        v = inArray(r, c)
        'formatting
        cell = IIf(IsNull(v), "", v)
        If TypeName(v) = "Date" Then cell = Format(v, fmtDate)
        'quote and escape
        If quoting = CSVUtilsQuote.All Or _
           (quoting = CSVUtilsQuote.NONNUMERIC And Not IsNumeric(v)) Or _
           InStr(cell, ",") > 0 Or InStr(cell, """") > 0 Or InStr(cell, vbCr) > 0 Or InStr(cell, vbLf) > 0 Then
          cell = Replace(cell, """", """""")
          cell = """" & cell & """"
        End If
        'add to array
        arrField(c) = cell
      Next
      arrRecord(r) = Join(arrField, ",") & recordSeparator
    Next
    If Err.Number <> 0 Then GoTo ErrorExit 'unexpected error
    
    ConvertArrayToCSV = Join(arrRecord, "")
    Exit Function
ErrorExit:
    ConvertArrayToCSV = ""
End Function


'
' ParseCSVToDictionary
'    return Dictionary whose key is value of keyColumn and whose value is a Collection of fields in the record
'
Public Function ParseCSVToDictionary(ByRef csvText As String, Optional ByRef keyColumn As Long = 1, Optional ByRef allowVariableNumOfFields As Boolean = False) As Object
    Dim coll As Collection
    Dim dict As Object
    Dim r As Long
    Set ParseCSVToDictionary = Nothing 'for error
    Set coll = ParseCSVToCollection(csvText, allowVariableNumOfFields)
    If coll Is Nothing Then Exit Function ' error
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 1 To coll.Count 'include header row
       Set dict(coll(r)(keyColumn)) = coll(r)
    Next
    Set ParseCSVToDictionary = dict
End Function

'
' GetFieldDictionary
'    return Dictionary whose key is field name and whose value is column number (1,2,3,...)  of the field
'
Public Function GetFieldDictionary(ByRef csvText As String) As Object
    Dim coll As Collection
    Dim c As Long
    Dim v
    Set coll = ParseCSVToCollection(csvText, True, True) 'parse header only
    Set GetFieldDictionary = Nothing ' for error
    If coll Is Nothing Then Exit Function ' Error
    Set GetFieldDictionary = CreateObject("Scripting.Dictionary")
    If coll.Count = 0 Then Exit Function ' no field (empty)
    For c = 1 To coll(1).Count
       v = coll(1)(c)
       GetFieldDictionary.Item(v) = c
    Next
End Function

' ------------- Private function/sub ---------------------------------------------------------------------

'
' find all separators in csvText
' - rowCount, colCount = size of array in csv
' - sepArray1234 = array of field info.  Their size => number of fields + 1.  Index start with 0, sepArray1234(number of fields) = 0
'                  sepArray1 = start pos of field, sepArray2 = field length, sepArray3 = nextSepType, sepArray4 = number of double quotes in field
Private Sub ParseCSV(ByRef rowCount As Long, ByRef colCount As Long, ByRef sepArray1() As Long, ByRef sepArray2() As Long, ByRef sepArray3() As Long, ByRef sepArray4() As Long, _
                     ByRef csvText As String, Optional ByRef allowVariableNumOfFields As Boolean = False)
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csvPos As Long
    Dim fieldText As String
    Dim nextSep As Long, nextSepType As Long, quoteCount As Long, fieldStart As Long, fieldLen As Long
    Dim colCountTmp As Long
    Dim sepIndex As Long, sepSize As Long
    
    sepSize = UBound(sepArray1)
    
    rowCount = 0
    colCount = 0 'max of colomn counts
    colCountTmp = 0 'current column count
    sepIndex = 0
    
    'for empty text
    If csvText = "" Then Exit Sub      'return empty collection
    
    ' Add trailing record separator if not
    If Right(csvText, 1) <> "" & vbCr And Right(csvText, 1) <> "" & vbLf Then
        csvText = csvText & vbCrLf
    End If
    
    'extract records and fields
    csvPos = 1
    Call FindNextSeparatorInit(csvText)
    Do While FindNextSeparator(csvText, csvPos, fieldStart, fieldLen, nextSepType, quoteCount)
        If Err.Number <> 0 Then Exit Do
        
        ' enhance array size if it is short
        If sepIndex + 1 > sepSize Then
            sepSize = sepSize * 2
            ReDim Preserve sepArray1(sepSize)  'new elements is initialized by 0
            ReDim Preserve sepArray2(sepSize)
            ReDim Preserve sepArray3(sepSize)
            ReDim Preserve sepArray4(sepSize)
        End If
        sepArray1(sepIndex) = fieldStart
        sepArray2(sepIndex) = fieldLen
        sepArray3(sepIndex) = nextSepType
        sepArray4(sepIndex) = quoteCount
        sepIndex = sepIndex + 1
        
        colCountTmp = colCountTmp + 1
        
        If nextSepType <> 1 Then ' next sep is record separator
            rowCount = rowCount + 1
            If colCount = 0 Then colCount = colCountTmp ' at initial row
            If Not allowVariableNumOfFields And colCount <> colCountTmp Then
                ErrorRaise 10001, "ParseCSVToCollection", "Syntax Error in CSV: numbers of fields are different among records"
                Exit Sub
            End If
            If colCountTmp > colCount Then colCount = colCountTmp
            colCountTmp = 0
        End If
    Loop
End Sub


' Find next separator (comma, CR, LF, CRLF) in inText starting with the position "start"
'   fieldStart = start position of found field
'   fieldLen   = length of found field
'   start = found separator + 1 (start of next field)
'   nextSepType = found separator type (1=comma, 2=CR or CRLF, 3=LF)
'   quoteCount = double quotation count in found field
'   return False if there is no next separator
'   * found field includes double quote (not yet parsing quotation syntax)
'   * assuming CR or LF exists at EOF
Private Sub FindNextSeparatorInit(ByRef inText As String)
    Dim lenText As Long
    lenText = Len(inText)
    nextSep1 = InStr(1, inText, ",")
    If nextSep1 = 0 Then nextSep1 = lenText + 1 'EOF
    nextSep2 = InStr(1, inText, "" & vbCr)
    If nextSep2 = 0 Then nextSep2 = lenText + 1 'EOF
    nextSep3 = InStr(1, inText, "" & vbLf)
    If nextSep3 = 0 Then nextSep3 = lenText + 1 'EOF
End Sub


Private Function FindNextSeparator(ByRef inText As String, _
                    ByRef start As Long, _
                    ByRef fieldStart As Long, _
                    ByRef fieldLen As Long, _
                    nextSepType As Long, ByRef quoteCount As Long) As Boolean
    Dim init_start As Long, lenText As Long
    Dim nextSep As Long, nextStart As Long
    
    FindNextSeparator = False
    
    lenText = Len(inText)
        
    If start > lenText Then Exit Function  'over run (no separator found in previous call)
        
    quoteCount = 0
    fieldStart = start
    
    Do While start <= lenText
        ' update nextSep(min of nextSep123), nextSepType, nextStart(next pos of next separator), nextSep123
        If nextSep1 < nextSep2 Then
            If nextSep1 < nextSep3 Then ' nextSep1 is smallest
                nextSep = nextSep1
                nextSepType = 1
                nextStart = nextSep + 1
                nextSep1 = InStr(nextStart, inText, ",")
                If nextSep1 = 0 Then nextSep1 = lenText + 1 'EOF
            Else ' nextSep3 is smallest
                nextSep = nextSep3
                nextSepType = 3
                nextStart = nextSep + 1
                nextSep3 = InStr(nextStart, inText, "" & vbLf)
                If nextSep3 = 0 Then nextSep3 = lenText + 1 'EOF
            End If
        Else
            If nextSep2 < nextSep3 Then ' nextSep2 is smallest
                nextSep = nextSep2
                nextSepType = 2
                nextStart = nextSep + 1
                If nextSep3 = nextSep2 + 1 Then ' CRLF
                    nextStart = nextStart + 1
                    nextSep3 = InStr(nextStart, inText, "" & vbLf)
                    If nextSep3 = 0 Then nextSep3 = lenText + 1 'EOF
                End If
                nextSep2 = InStr(nextStart, inText, "" & vbCr)
                If nextSep2 = 0 Then nextSep2 = lenText + 1 'EOF
            Else ' nextSep3 is smallest
                nextSep = nextSep3
                nextSepType = 3
                nextStart = nextSep + 1
                nextSep3 = InStr(nextStart, inText, "" & vbLf)
                If nextSep3 = 0 Then nextSep3 = lenText + 1 'EOF
            End If
        End If
        
        If nextSep > lenText Then  ' separator not found
            Exit Function
        End If
        
        Call StrCount(inText, start - 1, nextSep - 1, quoteCount) 'update number of double quates in [fieldStart, nextSep-1]
        start = nextStart
        
        If quoteCount Mod 2 = 0 Then  'if the number of double-quates is even, then the separator is not fake
            FindNextSeparator = True
            fieldLen = nextSep - fieldStart
            Exit Function
        End If
    Loop
    
    ErrorRaise 10002, "ParseCSVToCollection", "Syntax Error in CSV: illegal double-quote code"
End Function

'
' add number of double quotes in [n+1, p1] of Source to quoteCount
'
Private Sub StrCount(Source As String, n As Long, p1 As Long, ByRef quoteCount As Long)
    Dim ss As String
    Dim nn As Long
    Do
        ss = Mid(Source, n + 1, p1 - n) ' to avoid from feeding long string to InStr().
        nn = InStr(1, ss, """")
        If nn = 0 Then Exit Do
        n = n + nn
        quoteCount = quoteCount + 1
    Loop
End Sub

'
' Trim all before and after doube-quote
'   * text MUST include two or more double-quotes (")
Private Function TrimQuotes(ByRef text As String) As String
    'If InStr(text, """") = 0 Then Err.Raise 9999, "", "program error"
    Dim p0 As Long, p1 As Long
    Dim s As String
    
    'trim tail
    For p1 = Len(text) To 1 Step -1
      s = Mid(text, p1, 1)
      If (s = """") Then Exit For
    Next
    'trim head
    For p0 = 1 To p1
      s = Mid(text, p0, 1)
      If (s = """") Then Exit For
    Next
    'return
    TrimQuotes = Mid(text, p0 + 1, p1 - p0 - 1)
End Function



