# StartsWith

Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option.

```vb
Public Function StartsWith(ByVal Str As String, ByVal Value As String, ByVal StringComparison As VbCompareMethod) As Boolean
```

### Parameters

**Str** `String` <br>
A source string which will be examined.

**Value** `String` <br>
The string to compare

**ComparisonMethod** `VbCompareMethod` <br>
One of the enumeration values that specifies the rules for the search.

### Returns

`Boolean` <br>
`true` if this instance begins with `value`; otherwise, `false`.

### Exceptions

`ArgumentException` <br>
When `Str` is empty string.

`ArgumentOutOfRangeException` <br>
`ComparisonMethod` is not `vbTextCompare` or `vbBinaryCompare`.

## Examples

The following example finds the first occurrence of either letter a in the string.

```vb
Option Explicit
Public Sub Start()

    Debug.Print "Case insensitive search: " & StringH.StartsWith("The House of the Seven Gables", "the", VbCompareMethod.vbBinaryCompare)
    Debug.Print "Case sensitive search: " & StringH.StartsWith("The House of the Seven Gables", "the", VbCompareMethod.vbTextCompare)

End Sub

' Case insensitive search: False
' Case sensitive search: True
```

### Remarks

The StartsWith method compares the value parameter to the substring at the beginning of this string and returns a value that indicates whether they are equal.