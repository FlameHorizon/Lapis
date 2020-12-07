# StringExt.Substring Method

Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length. If length is not specified, selection continues to the end of the string.

```vb
Public Function Substring(ByVal Str As String, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As String
```

### Parameters

**Str** `String` <br>
A source string.

**StartIndex** `Long` <br>
The one-based starting character position of a substring in this instance.

**Length** `Long` <br>
Optional. The number of characters in the substring.

### Returns

`String` <br>
A string that is equivalent to the substring of length `length` that begins at `startIndex` in this instance, or `vbNullString` if `startIndex` is equal to the length of this instance and `length` is zero.

### Errors

`OnArgumentOutOfRange` <br>
`startIndex` plus `length` indicates a position not within this instance.

-or-

`startIndex` is less than one or `length` is less than zero.

## Examples

The following example illustrates a simple call to the Substring(String, Long, Long) method that extracts two characters from a string starting at the sixth character position (that is, at index six).

```vb
Option Explicit

Public Sub Start()

    Dim Str As String: Str = "This is a string"
    Dim StartIndex As Long: StartIndex = 6
    Dim Length As Long: Length = 2
    
    Dim SubStr As String
    SubStr = StringExt.Substring(Str, StartIndex, Length)
    Debug.Print SubStr

End Sub

' The example displays the following output:
'
' is
```

