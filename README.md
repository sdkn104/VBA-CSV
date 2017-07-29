# VBA-CSV

VBA Functions to parse CSV file

## The CSV File format

There is no definitive standard for CSV (Comma-separated values) file format, however the most commonly accepted definition is [RFC 4180](http://www.ietf.org/rfc/rfc4180.txt). VBA-CSV compliant with RFC 4180, while still allowing some flexibility where CSV files deviate from the definition.
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
