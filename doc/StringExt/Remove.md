# StringExt.Remove Method

Returns a new string in which all the characters in the current instance, beginning at a specified position and continuing through the last position, have been deleted.

```vb
Public Function Remove(ByVal Str As String, ByVal StartIndex As Long) As String
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**StartIndex** `Long` <br>
The zero-based position to begin deleting characters.

### Returns

`String` <br>
A new string that is equivalent to this string except for the removed characters.

### Errors

`OnArgumentOutOfRange` <br>
`StartIndex` is smaller than 0.

-or-

`StartIndex` is greater than the length of `Str`

## Examples

The following example demonstrates the Remove method. The next-to-last case removes all text starting from the specified index through the end of the string. The last case removes three characters starting from the specified index.

```vb
Option Explicit
Public Sub Start()

    Dim Str as String: Str = "abc---def"
    Debug.Print "1) " & Str
    Debug.Print "2) " & StringExt.Remove(Str, 3)

End Sub

' The example displays the following output:
' 1) abc---def
' 2) abc
```

