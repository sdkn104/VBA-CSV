Attribute VB_Name = "Examples"
'
' VBA-CSV
'
' Copyright (C) 2017- sdkn104 ( https://github.com/sdkn104/VBA-CSV/ )
'
' License MIT (http://www.opensource.org/licenses/mit-license.php)
'
Option Explicit


Sub tt()
'    Dim FSO As Object
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    With FSO.GetFile("C:\Users\sdkn1\Desktop\Book1.csv").OpenAsTextStream
'        csvText = .ReadAll
'        .Close
'    End With
'    Set FSO = Nothing
End Sub

Sub debugPrintCSV(csv As Variant)
    Dim r As Collection, f As Variant
    For Each r In csv
      Debug.Print "------"
      For Each f In r
        Debug.Print "[" & f & "]"
      Next
    Next
    Debug.Print "--------"

    Dim csva
    Dim i As Long, j As Long
    csva = ParseCSVToArray(csvText)
    
    For i = LBound(csva, 1) To UBound(csva, 1)
      For j = LBound(csva, 2) To UBound(csva, 2)
        Debug.Print "[" & csva(i, j) & "]"
      Next
    Next
End Sub

