VBA-CSV
==

VBA-CSV provides CSV parsers as VBA functions.

The CSV (Comma-Separated Values) parsers read text in CSV format and return Collection or Array of the table contents.
The parsers comply with [RFC4180](http://www.ietf.org/rfc/rfc4180.txt) of CSV format that allows commas, line breaks, and double-quotes included in the fields.

## Examples

`ParseCSVToCollection(csvText as String) As Collection`
```VB
Dim csv As Collection
Set csv = ParseCSVToCollection("aaa,bbb,ccc" & vbCrLf & "xxx,yyy,zzz")
' csv(1)(2)  --> bbb
' csv(2)(1)  --> xxx
```

a

```vb
bbb = 0
```

## Installation
 
1. Download the latest release
2. Import JsonConverter.bas into your project (Open VBA Editor, Alt + F11; File > Import File)
3. Add Dictionary reference/class
 - For Windows-only, include a reference to "Microsoft Scripting Runtime"
 - For Windows and Mac, include VBA-Dictionary

## The CSV File format

There is no definitive standard for CSV (Comma-separated values) file format, however the most commonly accepted definition is [RFC4180](http://www.ietf.org/rfc/rfc4180.txt). VBA-CSV compliant with RFC 4180, while still allowing some flexibility where CSV files deviate from the definition.
VBA-CSV accepts the CSV file that satisfies the following rules.

1.  Each record is located on a separate line, delimited by a line break (CRLF, *CR, or LF*).

       ```
       aaa,bbb,ccc CRLF
       zzz,yyy,xxx CRLF
       ```

2.  The last record in the file may or may not have an ending line break.

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
