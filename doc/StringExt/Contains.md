# Contains

Returns a value indicating whether a specified substring occurs within this string.

```vb
Public Function Contains(ByVal Str As String, ByVal Value As String) As Boolean
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**Value** `String` <br>
The string to seek.

### Returns

`Boolean` <br>
Indicates of `Value` substring exists in the `Str`.

## Examples

The following example finds the first occurrence of either letter a in the string.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "Does A11-003b contains letter B? " & StringH.Contains("A11-003b", "B")
    Debug.Print "Does A11-003b contains letter B?" & StringH.Contains("A11-003b", "a")

End Sub

' The example displays the following output:
' Does A11-003b contains letter B? True
' Does A11-003b contains letter z? False
```

### Remarks

This method performs an ordinal (case-sensitive and culture-insensitive) comparison. The search begins at the first character position of this string and continues through the last character position.