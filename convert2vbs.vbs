'
'Conversion Script
'
'  This script convert CSVUtils.bas to CSVUtils.vbs (VBScript version)
'  This script convert CSVUtils_Test.bas to CSVUtils_Test.vbs (VBScript version)
'

text = readFile("CSVUtils.bas")
call convert(text)
call writeFile("CSVUtils.vbs", text)

text = readFile("CSVUtils_Test.bas")
call convert(text)
call writeFile("CSVUtils_Test.vbs", text)

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "cmd.exe /c copy /B CSVUtils_Test.vbs+CSVUtils.vbs test.vbs"

'---------------------------------------------------------------------------------------

Sub convert(ByRef text)
text = ReplaceRE(text, "^(Attribute)", "'VBScript Version" & vbCrLf & "'  Error is always Fatal." & vbCrLf & "'  Array always starts with index 0" & vbCrLf & vbCrLf & "'$1")
text = ReplaceRE(text, "(Const IsVBA As Boolean =) True", "$1 False")
text = ReplaceRE(text, "(For Each rc In csv)", "For rr = 1 To recCnt : ri = rr-1 '$1")
text = ReplaceRE(text, "(For Each cc In rc)", "For ff = 1 To fldCnt : fi = ff-1 : cc = csv.Item(rr).Item(ff) '$1")
text = ReplaceRE(text, "(On Error Resume Next)", "'$1")
text = ReplaceRE(text, "(Resume Next)", "'$1")
text = ReplaceRE(text, "(\n|\r|\f)(Option Explicit)", "$1'$2")
text = ReplaceRE(text, "(\n|\r|\f)(\w+[:])", "$1'$2")
text = ReplaceRE(text, "(ReDim\s+\w+\s*)[(]\s*(\w+)\s* To \s*(\w+)\s*,\s*(\w+)\s* To \s*(\w+)", "$1($3-$2, $5-$4")
text = ReplaceRE(text, "(ReDim\s+\w+\s*)[(](.+) To (.+)[)]\s+As\s", "$1(($3)-($2)) As ")
text = ReplaceRE(text, "Optional\s+(ByRef|ByVal|)\s*(\w+)\s+As\s+\w+\s+=[^,)]+", "$2")
text = ReplaceRE(text, "As String", "")
text = ReplaceRE(text, "As Long", "")
text = ReplaceRE(text, "As Variant", "")
text = ReplaceRE(text, "As Object", "")
text = ReplaceRE(text, "As Boolean", "")
text = ReplaceRE(text, "As Collection", "")
text = ReplaceRE(text, "As Single", "")
text = ReplaceRE(text, "(\n|\r|\f)(.* GoTo .*)(\n|\r|\f)", "$1'$2$3")
text = ReplaceRE(text, "(\n|\r|\f)#", "$1'#")
text = ReplaceRE(text, "Debug.Print", "MsgBox")
text = ReplaceRE(text, "ConvertArrayToCSV[(]\s*(\w+)\s*[)]", "ConvertArrayToCSV($1,""yyyy/m/d"")")

End Sub



Function ReplaceRE(text, re, subst)
  Set regEx = New RegExp
  regEx.Pattern = re
  regEx.IgnoreCase = True
  regEx.Global = True
  ReplaceRE = regEx.Replace(text,subst)
End Function


Sub writeFile(fileName, text)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.CreateTextFile(fileName, True, False)
        .Write text
        .Close
    End With
End Sub

Function readFile(fileName)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.GetFile(fileName).OpenAsTextStream
        readFile = .ReadAll
        .Close
    End With
End Function



