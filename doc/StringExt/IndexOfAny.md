# StringExt.IndexOfAny Method

Reports the zero-based index of the first occurrence in this instance of any character in a specified array of characters.

```vb
Public Function IndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**AnyOf()** `String` <br>
Array of strings containing one or more characters to seek.

### Returns

`Long` <br>
The zero-based index position of the first occurrence in this instance where any character in `AnyOf` was found, -1 if no character in `AnyOf` was found.

### Errors

`OnArgumentError` <br>
`AnyOf` is not initialized.

## Examples

The following example finds the first occurrence of either 1 or 2 in the string.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringExt.IndexOfAny("Any 1 number is not 2.", ArrayH.StringArray("1", "2"))
End Sub

' The example displays the following output:
' 4
```

### Remarks

The search is case-insensitive. If value of `Str` parameter is empty, method returns -1.
