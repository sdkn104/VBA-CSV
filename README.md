VBA-CSV
=======

VBA-CSV provides CSV parsers as VBA functions.
The CSV (Comma-Separated Values) parsers read CSV text and return Collection or Array of the CSV table contents.
* The parsers are compliant with the CSV format defined in [RFC4180](http://www.ietf.org/rfc/rfc4180.txt), which allows commas, line
  breaks, and double-quotes included in the fields.


## Examples

#### ParseCSVToCollection(csvText as String) As Collection

```
Dim csv As Collection
Dim rec As Collection, fld As Variant
```

```
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

#### ParseCSVToArray(csvText as String) As Variant

    ```
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

#### Set(csvText as String) As Variant


## Installation
 
1. Download the latest release
2. Import CSVUtils.bas into your project (Open VBA Editor, Alt + F11; File > Import File)
3. Add reference to RegExp (VBScript ???)

## The CSV File format

There is no definitive standard for CSV (Comma-separated values) file format, however the most commonly accepted definition is 
[RFC4180](http://www.ietf.org/rfc/rfc4180.txt). VBA-CSV is compliant with RFC 4180, while still allowing some flexibility 
where CSV text deviate from the definition.
VBA-CSV accepts the CSV file that satisfies the following rules.

1.  Each record is located on a separate line, delimited by a line break (CRLF, *CR, or LF*).

       ```
       aaa,bbb,ccc CRLF
       zzz,yyy,xxx CRLF
       ```

2.  The last record in the file may or may not have an ending line break.
    *IF the last line is empty, the ending line break is necessary.*

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
     If fields are not enclosed with double quotes, then double quotes may not appear inside the fields.

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

8.    Spaces *(including tabs)* are considered part of a field and should not be ignored *(but some applications ignore them)*.
      *If fields are enclosed with double quotes, then leading and trailing spaces outside of double quotes are ignored.*

       ```
       " aaa", "bbb", ccc
       ```

## Author

[sdkn104](https://github.com/sdkn104)

## License

This software is released under the [MIT](https://opensource.org/licenses/mit-license.php) License.
