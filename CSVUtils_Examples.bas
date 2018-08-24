Attribute VB_Name = "CSVUtils_Examples"
'
' Examples for VBA-CSV
'
Option Explicit

'
' Example for ParseCSVToCollection()
'
Sub Example1()
    Dim csv As Collection
    Dim rec As Collection, fld As Variant

    Set csv = ParseCSVToCollection("aaa,bbb,ccc" & vbCr & "xxx,yyy,zzz")
    If csv Is Nothing Then
        Debug.Print Err.Number & " (" & Err.Source & ") " & Err.Description
    End If
    
    Debug.Print csv(1)(3) '----> ccc
    Debug.Print csv(2)(1) '----> xxx
    For Each rec In csv
      For Each fld In rec
        Debug.Print fld
      Next
    Next
End Sub

'
' Example for ParseCSVToArray()
'
Sub Example2()
    Dim csv As Variant
    Dim i As Long, j As Variant

    csv = ParseCSVToArray("aaa,bbb,ccc" & vbCr & "xxx,yyy,zzz")
    If IsNull(csv) Then
        Debug.Print Err.Number & " (" & Err.Source & ") " & Err.Description
    End If
    
    Debug.Print csv(1, 3) '----> ccc
    Debug.Print csv(2, 1) '----> xxx
    For i = LBound(csv, 1) To UBound(csv, 1)
      For j = LBound(csv, 2) To UBound(csv, 2)
        Debug.Print csv(i, j)
      Next
    Next
End Sub


'
' Example for ConvertArrayToCSV()
'
Sub Example3()
    Dim csv As String
    Dim a(1 To 2, 1 To 2) As Variant
    a(1, 1) = DateSerial(1900, 4, 14)
    a(1, 2) = "Exposition Universelle de Paris 1900"
    a(2, 1) = DateSerial(1970, 3, 15)
    a(2, 2) = "Japan World Exposition, Osaka 1970"
    
    csv = ConvertArrayToCSV(a, "yyyy/mm/dd")
    If Err.Number <> 0 Then
        Debug.Print Err.Number & " (" & Err.Source & ") " & Err.Description
    End If
    
    Debug.Print csv
End Sub

'
' Example for convert Excel Range to CSV, and writeFile(),
'             then readFile() and ParseCSV
'
Sub Example4()
    Dim text As String
    Dim csv As Variant
    Dim arr As Variant
        
    arr = ActiveSheet.Range("A1:C2")
    text = ConvertArrayToCSV(arr)
    Call writeFile("C:\Users\sdkn1\Desktop\Book1.csv", text)

    text = readFile("C:\Users\sdkn1\Desktop\Book1.csv")
    Set csv = ParseCSVToCollection(text)
    debugPrintResults csv
    csv = ParseCSVToArray(text)
    debugPrintResults csv
End Sub


'
' read text file and return String
'
Function readFile(Filename, Optional Encoding = "UTF-8") As String
    With CreateObject("ADODB.Stream")
        .Open
        .Charset = Encoding
        .LoadFromFile Filename
        readFile = .ReadText
        .Close
    End With
End Function

Function readFile2(Filename As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.GetFile(Filename).OpenAsTextStream
        readFile = .ReadAll
        .Close
    End With
End Function

'
' write text to file
'
Sub writeFile(fileName As String, text As String, Optional iomode As Long = 2)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FileExists(fileName) Then
      Call FSO.CreateTextFile(fileName, True, False)
    End If
    With FSO.OpenTextFile(fileName, iomode, TristateFalse) 'iomode:ForWriting(2),ForAppending(8);  format:TristateFalse(ASCII=ShiftJIS),TristateTrue(utf16)
        .Write text
        .Close
    End With
End Sub

Sub writeFile2(fileName As String, text As String, Optional Encoding As String = "UTF-8")
    With CreateObject("ADODB.Stream")
        .Mode = 3 'adModeReadWrite(3),...
        .Type = 2 'adTypeText(2), adTypeBinary(1)
        .Charset = Encoding '"UTF-8", "Shift_JIS", ...
        .Open
        .WriteText text, adWriteChar
        .SaveToFile fileName, 2 '2:adSaveCreateOverWrite
        .Close
    End With
End Sub


'
' Debug.Print the returned variable from the parser
'
Sub debugPrintResults(csv As Variant)
    
    Debug.Print "TypeName: " & TypeName(csv)
    If TypeName(csv) = "Collection" Then
        Dim r As Collection, f As Variant
        For Each r In csv
          Debug.Print "----------"
          For Each f In r
            Debug.Print "[" & f & "]"
          Next
        Next
        Debug.Print "--------"
    
    ElseIf TypeName(csv) = "String()" Then
        Dim i As Long, j As Long
        For i = LBound(csv, 1) To UBound(csv, 1)
          Debug.Print "----------"
          For j = LBound(csv, 2) To UBound(csv, 2)
            Debug.Print "[" & csv(i, j) & "]"
          Next
        Next
        Debug.Print "----------"
    
    Else
       Debug.Print "Not collection nor array"
    End If
End Sub
