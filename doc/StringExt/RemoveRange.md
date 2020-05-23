# StringExt.RemoveRange Method

Returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted.

```vb
Public Function RemoveRange(ByVal Str As String, ByVal StartIndex As Long, ByVal Count As Long) As String
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**StartIndex** `Long` <br>
The zero-based position to begin deleting characters.

**Count** `Long` <br>
The number of characters to delete.

### Returns

`String` <br>
A new string that is equivalent to this string except for the removed characters.

### Errors

`OnArgumentOutOfRange` <br>
`StartIndex` or `Count` is smaller than 0.

-or-

`StartIndex` plus `Count` is greater than the length of `Str`

## Examples

The following example demonstrates the RemoveRange method. The next-to-last case RemoveRanges all text starting from the specified index through the end of the string. The last case RemoveRanges three characters starting from the specified index.

```vb
Option Explicit
Public Sub Start()

    Dim Str As String: Str = "abc---def"
    Debug.Print "1) " & Str
    Debug.Print "2) " & StringExt.RemoveRange(Str, 3, 3)

End Sub

' The example displays the following output:
' 1) abc---def
' 2) abcdef
```

