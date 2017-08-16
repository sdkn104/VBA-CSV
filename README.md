VBA-CSV
=======

VBA-CSV provides CSV parsers and writer as VBA functions.
The CSV (Comma-Separated Values) parsers read CSV text and return Collection or Array of the CSV table contents. The CSV writer converts 2-dimensional array to CSV text.
* The parsers and writer are compliant with the CSV format defined in [RFC4180](http://www.ietf.org/rfc/rfc4180.txt), 
  which allows commas, line breaks, and double-quotes included in the fields.
* Test procedure and Examples are included.

## Examples

#### ParseCSVToCollection(csvText as String) As Collection

```vb.net
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
```

`ParseCSVToCollection()` returns a Collection of records each of which is a collection of fields.
If error occurs, it returns `Nothing` and the error information is set in `Err` object.

#### ParseCSVToArray(csvText as String) As Variant

```vb.net
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
```

`ParseCSVToArray()` returns a Variant that contains 2-dimensional array --- String(1 TO recordCount, 1 TO fieldCount).
If error occurs, it returns `Null` and the error information is set in `Err` object.
If input text is zero-length (""), it returns empty array --- String(0 TO -1).

#### ConvertArrayToCSV(inArray as Variant) As String

```vb.net
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
```


#### SetCSVUtilsAnyErrorIsFatal(value As Boolean)

```vb.net
    SetCSVUtilsAnyErrorIsFatal True
    SetCSVUtilsAnyErrorIsFatal False
```

This function changes error handling mode for CSV parsers and writer.

**False (default)** --- When run-time error occurs, the parser function returns special value (Nothing,  Null, etc.),
                    and the error information is set to properties of Err object.

**True**            --- Any run-time error that occurs is fatal (an error message is displayed and execution stops).

## Installation
 
1. Download the latest release
2. Import CSVUtils.bas into your project (Open VBA Editor, Alt + F11; File > Import File)
3. To enable RegExp, Add reference to "Microsoft VBScript Regular Expressions 5.5" (In VBA Editor, Tools > Reference Setting)

## Tested on

* Excel 2000
* Excel 2013

## The CSV File format

There is no definitive standard for CSV (Comma-separated values) file format, however the most commonly accepted definition is 
[RFC4180](http://www.ietf.org/rfc/rfc4180.txt). VBA-CSV is compliant with RFC 4180, while still allowing some flexibility 
where CSV text deviate from the definition.
The followings are the rules of CSV format such that VBA-CSV can handle correctly. 
(The rules indicated by *italic characters* don't exists in RFC4180)

1.  Each record is located on a separate line, delimited by a line break (CRLF, *CR, or LF*).

       ```
       aaa,bbb,ccc CRLF
       zzz,yyy,xxx CRLF
       ```

2.  The last record in the file may or may not have an ending line break.
    *The CSV file containing nothing (= "") is recognized as empty (it has no record nor fields).*

       ```
       aaa,bbb,ccc CRLF
       zzz,yyy,xxx
       ```

3.  Within each record, there may be one or more fields, separated by commas.
      
       ```
       aaa,bbb,ccc
       ```

4.  Each record should contain the same number of fields throughout the file.

5.  Each field may or may not be enclosed in double quotes.  

       ```
       "aaa","bbb","ccc" CRLF
       zzz,yyy,xxx
       ```
6.  Fields containing line breaks, double quotes, and commas should be enclosed in double-quotes.
       
       ```
       "aaa","b CRLF
       bb","ccc" CRLF
       zzz,yyy,xxx
       ```

7.  If double-quotes are used to enclose fields, then a double-quote
       appearing inside a field must be escaped by preceding it with
       another double quote.

       ```
       "aaa","b""bb","ccc"
       ```

8.    Spaces *(including tabs)* are considered part of a field and should not be ignored.
      *If fields are enclosed with double quotes, then leading and trailing spaces outside of double quotes are ignored.*

       ```
       " aaa", "bbb", ccc
       ```

## Author

[sdkn104](https://github.com/sdkn104)

## License

This software is released under the [MIT](https://opensource.org/licenses/mit-license.php) License.
