# IndexOf

Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a array.

```vb
Public Function LastIndexOfAny(ByVal Str As String, ByRef AnyOf() As String) As Long
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**AnyOf** `String()` <br>
Array of strings containing one or more characters to seek.

### Returns

`Long` <br>
The zero-based index position of the last occurrence in this instance where any character in `AnyOf` was found, -1 if no character in `AnyOf` was found.

### Exceptions

`ArgumentException` <br>
When `Str` is empty string.

`ArgumentException` <br>
`AnyOf` is not initialized.

## Examples

The following example finds the first occurrence of either 1 or 2 in the string.

```vb
Option Explicit
Public Sub Start()
    Debug.Print StringH.LastIndexOfAny("Any 1 number is not 2. but it's every 1.", ArrayH.StringArray("1", "2"))
End Sub

' The example displays the following output:
' 38
```

### Remarks

The search is case-insensitive. If value of `Str` parameter is empty, method returns -1.